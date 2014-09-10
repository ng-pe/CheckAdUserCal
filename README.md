CheckAdUserCal
==============

CheckAdUserCal : Terminal Server User Cal Monitoring for NSClient++

Overview
--------
CheckAdUserCal is a `command-line utility`_ that help you to monitor Microsoft Terminal Server/RDS User CAL.
You can use CheckAdUserCal with NSClient++ or standalone.
Coded with Microsoft VisualC# Express Edition 2010


Features
--------
- Calculate Free License 
- Calculate Registred license

Screenshot
----------

CheckAdUserCal v1.1 (console Mode)
Check Terminal Server User CAL on Active Directory
INFO > Adding TSCal Server : lxxx01
INFO > Get Data from ActiveDirectory... ( LDAP://DC=XXXX,DC=ADDS )
+ Active Directory Terminal Server User License :
         - Cal Used      = 1003
         - Cal Expired   = 3
         - Cal Registred = 1250

+ Cal Server Allocation
  LS Server : lvsad01
         - license issued = 1006
         - license used = 1003
         - registered license = 1250
         - license free = 247

+ Global Cal License Count = 247 free / 1250

+ Delivered license by 'UNKNOWN' LS = 0


Press any key to exit!