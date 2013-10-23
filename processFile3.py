
"""
Description:
This script was written to process a specific excel spreadsheet and go to the directory and pull an  email address 
for of the entries from directory.utexas.edu.

"""

import ldap
import xlrd
import operator

server = 'ldap://directory.utexas.edu'

wb = xlrd.open_workbook('JimFaculty.xls')
#wb = xlrd.open_workbook('TestSheet2.xls')
wb.sheet_names()
sh = wb.sheet_by_index(0)

con = ldap.initialize(server)
con.simple_bind()
base_dn = 'ou=people,dc=directory,dc=utexas,dc=edu'
attrs = ['cn','uid','title','mail','ou', 'utexasEduPersonPubAffiliation']
filterList=[]
d = {}

f = open('DataForJimFacultyStaffAndStudent', 'w')

# build a dictionary of faculty names as keys and rooms as values
for rownum in range(sh.nrows):
    list1 = sh.row_values(rownum)
    zero,four,five,nine,ten=operator.itemgetter(0,4,5,9,10)(list1)
    if not list1[0] in d:
        d[list1[0]] = []
    list2=(zero,four,five,nine,ten)
    if not four.isspace():
        if not (four,five) in d[list2[0]]:
            dictentry = [(four,five)]
            d[list2[0]] = d[list2[0]] + dictentry
    if not nine.isspace():
        if not (nine,ten) in d[list2[0]]:
            dictentry = [(nine,ten)]
            d[list2[0]] = d[list2[0]] + dictentry


# this cleans up blank lines 
for item in d.items():
    if item[0].isspace():
        del d[item[0]]

# move all the directory stuff here 

for key in d:
    fullName=key
    fullNameList=fullName.split(',')
    lastName=fullNameList[0].strip()
    if len(fullName.split(',')) > 1:
        firstName=''
        firstName=fullName.split(',')[1].lstrip().split(' ')[0]
        filter='(&(sn=%(last)s)(givenName=%(first)s)(utexasEduPersonPubAffiliation=faculty))' % {"last": lastName, "first": firstName}

    result = con.search_s( base_dn, ldap.SCOPE_SUBTREE, filter, attrs )
#    print result
    # result is a list(tuple(dict))
    email=''
    if len(result) > 0:
        for i in range(len(result)):
            if 'mail' in result[i][1]:
                email=result[i][1]['mail'][0]
            else:
                email='none available'
        dictemail=[(email)]
        if not email in d[key]:
            d[key] = d[key] + dictemail

        dictemail=[]

    # if not faculty try staff
    if len(result) == 0:
        filter='(&(sn=%(last)s)(givenName=%(first)s)(utexasEduPersonPubAffiliation=staff))' % {"last": lastName, "first": firstName} 
        result = con.search_s( base_dn, ldap.SCOPE_SUBTREE, filter, attrs )
        if len(result) > 0:
            for i in range(len(result)):
                if 'mail' in result[i][1]:
                    email=result[i][1]['mail'][0]
                else:
                    email='none available'
            dictemail=[(email)]
            if not email in d[key]:
                d[key] = d[key] + dictemail

            dictemail=[]

    # if not faculty or staff try student
    if len(result) == 0:
        filter='(&(sn=%(last)s)(givenName=%(first)s)(utexasEduPersonPubAffiliation=student))' % {"last": lastName, "first": firstName}
        result = con.search_s( base_dn, ldap.SCOPE_SUBTREE, filter, attrs )
        if len(result) > 0:
            for i in range(len(result)):
                if 'mail' in result[i][1]:
                    email=result[i][1]['mail'][0]
                else:
                    email='none available'
            dictemail=[(email)]
            if not email in d[key]:
                d[key] = d[key] + dictemail

            dictemail=[]

# put email first in list
for key in d:
     d[key].sort()

for key in d:
    f.write("%s" % key)
    for item in d[key]:
        x = item
        f.write("\t")
        f.write("|")
        for y in x:
            str(y)
            f.write("%s" % y)
    f.write("\n")

f.close()

