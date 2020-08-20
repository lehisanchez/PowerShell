# --------------------------
# Batch Add Members to Group
# --------------------------

# Variables: 
# FILENAME
# GROUPNAME
# COLUMNHEADER

Import-CSV "C:\FILENAME.csv" | % {Add-ADGroupMember -Identity GROUPNAME -Member $_.COLUMNHEADER}