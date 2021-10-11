# Artemius-Outlook-Additional-Information - (AOAI)

# Description

This add-on extends the MS Outlook user interface and adds additional features to inform users when working with external incoming and outgoing email. With this add-on, you will be able to tag external emails in MS Outlook. You will also be able to inform users if their emails are sent outside the corporate network.

# Work architecture

The add-on uses VSTO (COM) technology to integrating with MS Outlook

## Author

[_KUL](https://github.com/isKUL)

# Add-on modules and their features

- Module for marking emails with MS Outlook color categories.

The list of trusted domain names is analyzed. The configuration of processing incoming and outgoing emails is analyzed. Events for receiving emails or viewing emails are enabled. If the domain of the sender of the email message is not included in the trusted domain, then the message is marked with a colored category that the email is external.

![marking](https://github.com/isKUL/AOAI/blob/main/_img/marking.png?raw=true)

- Notification module when sending emails.

The list of trusted domain names is analyzed. The configuration of notification processing is analyzed. Events are included for the analysis of emails when sending. If the outgoing email contains an addressee who is not a member of the trusted domain, then an information window is displayed to the user. In the information window, you can set the required warning - remind about the company's policy and polite communication with external contractors, remind users about trade secrets and the prohibition of the transfer of classified information, etc.

![notification](https://github.com/isKUL/AOAI/blob/main/_img/notification.png?raw=true)
	
# Configuration modes

The add-on has a common configuration file, thanks to this, a mechanism for obtaining configurations in different ways is provided. At the moment, the following methods of obtaining the configuration are implemented:

## Local

The configuration is stored directly inside the program variable. All values are set before compilation, then after compilation the program uses the built-in configuration.

## Centralized using Active Directory.

The user is a member of an Active Directory domain and wishes to use the AOAI add-on. It is possible to centrally distribute the configuration through Active Directory by saving the configuration in the `houseIdentifier` attribute of a `contact` class object named `*AOAI*`. Thus, all domain users when starting MS Outlook will download the AOAI add-on, and the add-on will receive the configuration from the Active Directory service object.

## Configuration load priority

From high to low: Centralized using Active Directory, Local

# Distribution

For the convenience of users, you can create an MSI installation package to install the add-on on a computer for all users and register the add-on keys in the MS Outlook registry. The main feature is that you must create a **trusted certificate** with which you **must sign** the **manifest** and the **MSI** so that the users' computers trust the add-on. For example, there is an installation project with a self-signed certificate.

# Detailed configuration
...
