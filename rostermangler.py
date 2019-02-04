#!/usr/bin/env python3
"""Mangle roster data from California State 4-H website and Mail chimp."""

import sys
import os
import csv
import pyexcel_ods
import collections
import re
import argparse


class Person:
    "Structure for data about a single member or volunteer"
    MIN_AGE = 5

    def __init__(self, first_name="", last_name="", last_name_first_name=None, phone="", email="", age=0, \
                 nickname=None, role=None, address=None, city=None):
        "Initalize structure"
        if last_name_first_name:
            self.last_name, self.first_name = last_name_first_name.split(', ')
        else:
            self.first_name = first_name
            self.last_name = last_name
        self.phone = phone
        self.email = email.lower()
        self.age = age
        self.nickname = nickname
        self.role = role
        self.address = address
        self.city = city

    def __repr__(self):
        "String presentation of class"
        return f"Person({self.first_name!r}, {self.last_name!r}, phone={self.phone!r}, email={self.email!r}, " \
               f"age={self.age!r}, nickname={self.nickname!r}, role={self.role!r}, address={self.address!r}, " \
               f"city={self.city!r})"

    def __str__(self):
        "Pretty print representation"
        if self.nickname:
            pretty = f"{self.first_name} ({self.nickname}) {self.last_name}"
        else:
            pretty = f"{self.first_name} {self.last_name}{os.linesep}"
        if self.role:
            pretty += f"    {self.role}{os.linesep}"
        if self.phone:
            pretty += f"    Phone: {self.phone}{os.linesep}"
        if self.email:
            pretty += f"    Email: {self.email}{os.linesep}"
        if self.address:
            pretty += f"    Address: {self.address}{os.linesep}"
        if self.city:
            pretty += f"    City: {self.city}{os.linesep}"
        if self.age >= self.MIN_AGE:
            pretty += f"    Age: {self.age}{os.linesep}"
        return pretty

    def __eq__(self, other):
        "Check for equality"
        return self.first_name.lower() == other.first_name.lower() and self.last_name.lower() == other.last_name.lower()

    @property
    def valid(self):
        "Check if this person is real or just blank"
        return self.first_name and self.last_name

    @staticmethod
    def _update_attr(this, other, field, resolve=None):
        "Update a field if no conflicts"
        if getattr(this, field) and getattr(other, field) and getattr(this, field) != getattr(other, field):
            if resolve == "concat":
                setattr(this, field, ", ".join((getattr(this, field), getattr(other, field))))
            elif resolve == "replace":
                setattr(this, field, getattr(other, field))
            else:
                raise ValueError(f"Update conflict {field}: this={getattr(this, field)!r}, other={getattr(other, field)!r}")
        else:
            setattr(this, field, getattr(other, field))

    def update(self, other):
        "Update this object with missing information from another one"
        if self != other:
            raise ValueError(f"Attempting to update \"{self.first_name} {self.last_name}\" from "\
                             f"\"{other.first_name} {other.last_name}\"")
        else:
            self._update_attr(self, other, "phone")
            self._update_attr(self, other, "email", "replace")
            self._update_attr(self, other, "nickname")
            self._update_attr(self, other, "role", "concat")
            self._update_attr(self, other, "address")
            self._update_attr(self, other, "city")

class Family:
    "Representation of a family group"

    def __init__(self, parents=[], children=[]):
        "Initalize blank family"
        self.parents = [p for p in parents if p.valid]
        self.children = [c for c in children if c.valid]

    def __repr__(self):
        "(non constructable) representation"
        return f"Family({os.linesep}" \
               f"    parents={self.parents!r},{os.linesep}" \
               f"    children={self.children!r}{os.linesep}" \
               ")"

    @property
    def last_names(self):
        "A set of all last names in the family"
        return {p.last_name for p in self.parents + self.children}

    @property
    def family_name(self):
        "A representation of the name of the family as a whole"
        last_names = list(self.last_names)
        last_names.sort()
        for name_a in last_names:
            for name_b in last_names:
                if name_a != name_b and (name_a.startswith(name_b) or name_a.endswith(name_b)):
                    return name_a + " family"
        return " and ".join(last_names) + " family"

    @property
    def individual(self):
        "Returns true if this family only has one person in it"
        if len(self.parents) == 1 and not self.children:
            return self.parents[0]
        if len(self.children) == 1 and not self.parents:
            return self.children[0]
        return None

    @property
    def family_phone(self):
        "Returns common family phone number"
        phones = collections.Counter([p.phone for p in self.parents + self.children])
        return [phone for phone, count in phones.most_common(2) if count > 1 and phone]

    @property
    def family_email(self):
        "Return common family email address if there is one"
        emails = collections.Counter([p.email for p in self.parents + self.children])
        return [email for email, count in emails.most_common(2) if count > 1 and email]

    @property
    def family_address(self):
        "Return common family address"
        if self.children:
            return (self.children[0].address, self.children[0].city)
        return (self.parents[0].city,)

    def sort(self, key=lambda p: (p.last_name, p.first_name)):
        "Trigger a sort on the parent and child lists. Default alphabetical by first name"
        self.parents.sort(key=key)
        self.children.sort(key=key)

    @staticmethod
    def _add_person(group, new_person):
        "Adds a new person to a list if not already present"
        if new_person.valid:
            for person in group:
                if new_person == person:
                    person.update(new_person)
                    return
            group.append(new_person)

    def add_or_update_parent(self, new_parent):
        "Adds a new parent to the family if not already present"
        self._add_person(self.parents, new_parent)

    def add_or_update_child(self, new_child):
        "Adds a new child to the family if not already present"
        self._add_person(self.children, new_child)

    def has_parent(self, first_name, last_name):
        "Check if this family has an adult with a given first and last name"
        for individual in self.parents:
            if individual.last_name == last_name and individual.first_name == first_name:
                return True
        return False

    def has_child(self, first_name, last_name):
        "Check if this family has a child with a given first and last name"
        for individual in self.children:
            if individual.last_name == last_name and individual.first_name == first_name:
                return True
        return False


def get_adult_volunteers_as_people(sheet, keys):
    "Return a list of People structures from Adult Volunteers sheet"
    return [Person(last_name_first_name=row[0], email=row[1], role=row[2], city=row[5].split(',')[0]) \
            for row in sheet if len(row) > 5]

def get_members_as_families(sheet, keys):
    "Return a list of Family data structures from Members sheet"
    families = []
    for row in sheet:
        if len(row) < 12:
            continue
        member = Person(last_name_first_name=row[0], email=row[1], phone=row[2], address=row[3], city=row[4],
                        age=int(row[12]))
        parent1 = Person(row[5], row[6], phone=row[7])
        parent2 = Person(row[8], row[9], phone=row[10], email=row[11])
        for family in families:
            if family.has_parent(parent1.first_name, parent1.last_name) or \
               family.has_parent(parent2.first_name, parent2.last_name):
                family.add_or_update_child(member)
                family.add_or_update_parent(parent1)
                family.add_or_update_parent(parent2)
                break
        else:
            families.append(Family([parent1, parent2], [member]))
    for fam in families:
        fam.sort()
    return families


def get_families_from_ucnar_ods(ods_file):
    "Convert an UCNAR export ODS file into a list of Family data structures"
    # Crack open the workbook file
    book = pyexcel_ods.get_data(ods_file)
    members_sheet = book['Members']
    member_keys = members_sheet.pop(0)
    adults_sheet = book['Adult Volunteers']
    adult_keys = adults_sheet.pop(0)
    # Parse the members sheet
    families = get_members_as_families(members_sheet, member_keys)
    # Parse the adult voluteers sheet
    adults = get_adult_volunteers_as_people(adults_sheet, adult_keys)
    # Unify the results
    for adult in adults:
        for family in families:
            if family.has_parent(adult.first_name, adult.last_name):
                family.add_or_update_parent(adult)
                break
        else:
            families.append(Family(parents=[adult]))
    return families

def get_members_and_volunteers(ods_file):
    "Get members and adult volunteers list from ODS excport."
    sheet = pyexcel_ods.get_data(ods_file)
    members = sheet['Members']
    member_keys = members.pop(0)
    members_email_dict = {row[1]: row for row in members if row}
    adults = sheet['Adult Volunteers']
    adult_keys = adults.pop(0)
    adults_email_dict = {row[1]: row for row in members if row}
    return member_keys, members_email_dict, adult_keys, adults_email_dict


def get_mailchip_data(mailchimp_csv):
    "Load CSV file from mailchimp export"
    sheet = list(csv.reader(mailchimp_csv, delimiter=","))
    sheet_keys = sheet.pop(0)
    email_dict = {row[0]: row for row in sheet}
    return sheet_keys, email_dict


def extra_in_mailchimp(mailchimp_email_dict, members_email_dict, adults_email_dict):
    "Find the email rows that are extra in mailchimp"
    known_emails = set(members_email_dict.keys())
    known_emails.update(adults_email_dict.keys())
    print(known_emails)
    unknown = []
    for email, row in mailchimp_email_dict.items():
        if email not in known_emails:
            unknown.append(row)
    return unknown


def missing_from_mailchimp(mailchimp_email_dict, members_email_dict, adults_email_dict):
    "Find emails that aren't in mailchimp"
    missing_members = []
    for email, row in members_email_dict.items():
        if email not in mailchimp_email_dict.keys():
            missing_members.append(row)
    missing_adults = []
    for email, row in adults_email_dict.items():
        if email not in mailchimp_email_dict.keys():
            missing_adults.append(row)
    return missing_members, missing_adults


def roster_merge(roster_input, mailchimp_export):
    "Program entry"
    # Parse inputs
    member_keys, members_email_dict, adult_keys, adults_email_dict = get_members_and_volunteers(roster_input)
    mailchimp_sheet_keys, mailchimp_email_dict = get_mailchip_data(mailchimp_export)
    # Make possible remove output
    possible_rm = extra_in_mailchimp(mailchimp_email_dict, members_email_dict, adults_email_dict)
    with open("possible_remove.csv", "wt") as possible_rm_file:
        writer = csv.writer(possible_rm_file, delimiter=",")
        for row in possible_rm:
            writer.writerow(row)
    # Make possible add output
    possible_add_members, possible_add_adults = missing_from_mailchimp(mailchimp_email_dict,
                                                                       members_email_dict,
                                                                       adults_email_dict)
    with open("possible_add.csv", "wt") as possible_add_file:
        writer = csv.writer(possible_add_file, delimiter=",")
        for row in possible_add_members:
            writer.writerow(row)
        for row in possible_add_adults:
            writer.writerow(row)


def print_table_row(heading, value, l_padding=" "*4):
    "Prints something in HTML table row brackets"
    if value:
        print(f"{l_padding}<tr><th>{heading}</th><td>{value}</td></tr>")


def roster(ods_file_name):
    "Make a pretty roster out of the state 4-H Export"
    families = get_families_from_ucnar_ods(ods_file_name)
    families.sort(key=lambda fam: fam.family_name)
    for fam in families:
        print('<div class="roster_family" id="{0}"><a name="{0}"></a>'.format(fam.family_name))
        indiv = fam.individual
        if indiv:
            print(f"  <h3>{indiv.last_name}, {indiv.first_name}</h3>")
            if indiv.role:
                print(f"  <h4>{indiv.role}</h4>")
            print("  <table>")
            print_table_row("Email", indiv.email)
            print_table_row("Phone", indiv.phone)
            print_table_row("Address", indiv.address)
            print_table_row("City", indiv.city)
            print("  </table>")
        else:
            print(f"  <h3>{fam.family_name}</h3>")
            print("  <table>")
            for email in fam.family_email:
                print_table_row("Email", email)
            for phone in fam.family_phone:
                print_table_row("Phone", phone)
            print_table_row("Address", ", ".join(fam.family_address) if fam.family_address else None)
            print("  </table>")
            print("  <table>")
            print("    <tr><th>Adults</th><th>Children</th><td>")
            print("      <tr><td><table>")
            for adult in fam.parents:
                print(f'        <tr><th colspan="2">{adult.first_name} {adult.last_name}</th></tr>')
                if adult.role:
                    print(f'        <tr><td colspan="2">{adult.role}</td></tr>')
                print_table_row("Email", adult.email if adult.email not in fam.family_email else None, " "*8)
                print_table_row("Phone", adult.phone if adult.phone not in fam.family_phone else None, " "*8)
            print("      </table></td><td><table>")
            for child in fam.children:
                print(f'        <tr><th colspan="2">{child.first_name} {child.last_name}</th></tr>')
                print_table_row("Email", child.email if child.email not in fam.family_email else None, " "*8)
                print_table_row("Phone", child.phone if child.phone not in fam.family_phone else None, " "*8)
            print("      </table></td></tr>")
            print("  </table>")
        print("</div>")

def main():
    "Program entry point"
    parser = argparse.ArgumentParser()
    parser.add_argument("-m", "--merge", nargs=2, help="Merge state 4-H export and Mailchimp Export")
    parser.add_argument("-r", "--roster", help="Make a pretty roster out of the state 4-H Export")
    args = parser.parse_args()
    if args.merge:
        roster_merge(args.merge[0], open(args.merge[1], 'rt'))
    if args.roster:
        roster(args.roster)


if __name__ == '__main__':
    main()
