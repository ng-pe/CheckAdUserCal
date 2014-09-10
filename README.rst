CheckAdUserCal
==============

CheckAdUserCal : Terminal Server User Cal Monitoring for NSClient++ and standalone

Overview
--------
CheckAdUserCal is a `command-line utility` that help you to monitor Microsoft Terminal Server/RDS User CAL.
You can use CheckAdUserCal with NSClient++ or standalone.

Coded with Microsoft VisualC# Express Edition 2010


Features
--------
- Show Used, Expired and registred TS CAL License 
- Show user that have an CAL registred
- Use with NSClient++ to report/alerts in NAGIOS ;)


HowTo use with NSClient++
-------------------------

Registre the External Scripts command line in NSC.ini
.. sourcecode:: console
[External Scripts]
check_tscal=scripts\TSCal_Monitoring.exe -s tscalsrv01 -w 15 -l "LDAP://DC=DOMAIN,DC=ADDS"

In Nagios/Centreon use "check_nrpe -H $HOSTADDRESS$ -t 60 -c check_tscal" command to check 


Screenshot
----------
.. sourcecode:: console

CheckAdUserCal v1.1 (console Mode)
Check Terminal Server User CAL on Active Directory
INFO > Adding TSCal Server : tscalsrv01
INFO > Get Data from ActiveDirectory... ( LDAP://DC=DOMAIN,DC=ADDS )
+ Active Directory Terminal Server User License :
         - Cal Used      = 1003
         - Cal Expired   = 3
         - Cal Registred = 1250

+ Cal Server Allocation
  LS Server : tscalsrv01
         - license issued = 1006
         - license used = 1003
         - registered license = 1250
         - license free = 247

+ Global Cal License Count = 247 free / 1250

+ Delivered license by 'UNKNOWN' LS = 0


Press any key to exit!