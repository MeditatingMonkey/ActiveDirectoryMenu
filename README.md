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
  
Screenshot's attached below.
  
  1st Window, search the name of the user.
  
  ![Screenshot 2023-06-04 181038](https://github.com/MeditatingMonkey/Directory_Auditor./assets/68747956/602184ac-7f29-42fd-af05-18f732b0cd6b)

  Select the Name and hit Submit.
  
  ![Screenshot 2023-06-04 181206](https://github.com/MeditatingMonkey/Directory_Auditor./assets/68747956/f5250662-3c61-4388-a0ce-beb5ef1d5810)

  ![Screenshot 2023-06-04 181320](https://github.com/MeditatingMonkey/Directory_Auditor./assets/68747956/39faa782-d931-4102-aa45-fbe8d41391d6)
  
  ![Screenshot 2023-06-04 181409](https://github.com/MeditatingMonkey/Directory_Auditor./assets/68747956/8baad7e9-bff9-41e7-93f7-bd74fd7ba3d3)
  
  ![Screenshot 2023-06-04 181429](https://github.com/MeditatingMonkey/Directory_Auditor./assets/68747956/de24f01a-e50a-426c-a729-784daa72482c)
    
  ![Screenshot 2023-06-04 181524](https://github.com/MeditatingMonkey/Directory_Auditor./assets/68747956/83d48bda-7c8f-4822-babb-cb87096a9164)
  
  ![Screenshot 2023-06-04 181550](https://github.com/MeditatingMonkey/Directory_Auditor./assets/68747956/c8fbf458-931a-4ede-8219-648450491376)
  
  ![Screenshot 2023-06-04 181704](https://github.com/MeditatingMonkey/Directory_Auditor./assets/68747956/9ea09dc5-59d2-4b6d-b9d1-1e307d4e9176)
  
  ![Screenshot 2023-06-04 181730](https://github.com/MeditatingMonkey/Directory_Auditor./assets/68747956/ef486e5f-c169-45db-b109-c4bb0abe553c)
  
  ![Screenshot 2023-06-04 181822](https://github.com/MeditatingMonkey/Directory_Auditor./assets/68747956/052bba8c-ab34-4d8b-9ee1-8598b47cbffa)
  
  ![Screenshot 2023-06-04 181841](https://github.com/MeditatingMonkey/Directory_Auditor./assets/68747956/db6b54b2-308f-4ef9-b7ef-7d003635ecd5)
  
  ![Screenshot 2023-06-04 181900](https://github.com/MeditatingMonkey/Directory_Auditor./assets/68747956/4f8b62d6-bae8-4dc1-b9cb-4fdf843cea65)
  
  ![Screenshot 2023-06-04 181917](https://github.com/MeditatingMonkey/Directory_Auditor./assets/68747956/54c800bf-c5ab-4076-97f6-ea247ab49a53)
  
  ![Screenshot 2023-06-04 182046](https://github.com/MeditatingMonkey/Directory_Auditor./assets/68747956/c9344137-04c1-42e4-96dc-2a4581215c05)







