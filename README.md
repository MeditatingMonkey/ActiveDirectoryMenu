# <h2> **ADMenu**  
  
ADMenu is a PowerShell script designed to simplify and streamline tasks related to user management in an Active Directory (AD) environment. This project incorporates a Graphical User Interface (GUI) that provides administrators with an intuitive and user-friendly way to perform common AD tasks. The focus of this tool is to enhance efficiency, reduce complexity and minimize the potential for errors that can occur when managing users within AD.

This tool provides the following options to administrators:

**Unlock Account:** This feature allows administrators to unlock user accounts that have been locked due to incorrect password entries or other security reasons.

**Close User:** This feature enables the administrator to disable a user account, preventing its further usage. This function is often used when an employee leaves the company or changes roles.

**Remote in the Computer:** This option allows the administrator to remotely access user computers for troubleshooting or support purposes.

**Group Access:** With this feature, an administrator can add users to specific security groups based on their roles and responsibilities. Once added, the system automatically sends an email to the user notifying them about their new access.

**Software Deployment:** This function uses System Center Configuration Manager (SCCM) to pull information about a user's computer. It then adds the computer to the appropriate security group for software deployment. An email notification is sent to the user once the process is complete.

**Reset Password:** This option provides a straightforward way to reset user passwords in case of forgetfulness or security breaches.

**User Details:** This feature allows the administrator to view all relevant details of a user account. This includes, but is not limited to, full name, user role, last login time, etc.

# <h2> **Instructions**
1) To use this download all the Powershell files and update all the Organization Units as per your organization.
2) Since "The Menu.ps1" calls all the powershell files fix all the addresses of the ps files in the "The Menu.ps1" file.
3) Check the email address too in the "Application Deployment.ps1" and "Group Deployment.ps1"

# <h2> **Conclusion and Future aspects**
ActiveDirectoryMenu not only simplifies the overall process of user management in AD, but it also creates a more controlled environment by sending out email notifications whenever changes are made. This project aims to increase productivity, optimize workflow, and enhance the overall security of your AD management.

Future developments of this project may include more advanced features like a user audit trail, batch processing for large user groups, and enhanced customization of the user interface.

Feel free to contribute to the project, open issues, and submit pull requests. Your feedback and ideas can go a long way in improving and expanding the functionality of this tool.




