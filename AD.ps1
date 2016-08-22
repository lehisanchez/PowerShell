# Active Directory Scripts

# Batch Add Members to Group
# Variables: file.csv,GROUPNAME,COLUMNHEADER
Import-CSV "C:\file.csv" | % {Add-ADGroupMember -Identity GROUPNAME -Member $_.COLUMNHEADER}