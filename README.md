---

# 🤖 Automated Referral System

### **Project Summary**

A comprehensive Google Apps Script solution developed to automate and manage a referral program for a ride-sharing company. This system streamlines the end-to-end referral process, from data validation and eligibility checks to automated email communication and performance tracking. By replacing manual workflows with a robust, automated system, this project significantly improves operational efficiency, reduces human error, and ensures timely communication with drivers.

---

### **Key Features**

* **Real-time Data Validation:** Automatically validates incoming referral submissions against multiple datasets (e.g., "Scouts" and "Churned Users") to determine eligibility.
* **Intelligent Email Automation:** Sends personalized, automated emails to both the referring driver ("scout") and the referred driver based on specific eligibility and status criteria. This includes confirmation, non-confirmation, and bonus-related updates.
* **Duplicate Detection:** Includes a sophisticated duplicate check to prevent the same driver from being referred multiple times, ensuring fairness and data integrity.
* **Automated Data Management:** Segregates and organizes referral data into different Google Sheets based on validation status (e.g., "Valid Referrals," "Not Eligible Referrals," "Compensation Due").
* **Performance Tracking:** Automatically updates driver activity and status, enabling the program to track key metrics like trip completion and reactivation dates.
* **Scalable & Low-Maintenance:** Built on Google Apps Script, the solution is cloud-based, requiring no infrastructure setup. It can be easily integrated with Google Forms and Google Sheets, making it simple to manage and scale as the program grows.

---

### **How It Works**

The system operates through a series of interconnected functions triggered by new form submissions or time-based triggers. 

1.  **Form Submission:** A new referral form entry is submitted to the "Referrals" sheet.
2.  **Validation Check:** The `validateNewRecords` function instantly checks the scout's and referred driver's details against internal "Scouts" and "Churned Users" sheets.
3.  **Eligibility & Duplication:** The `checkLastPhoneNumberDuplicate` function verifies if the referred driver has been submitted before.
4.  **Conditional Logic:** The `checkAndSendEmails` function uses the validation results to determine the appropriate course of action. It routes data to the correct sheet (`Valid Referrals` or `Not Eligible Referrals`) and triggers the right email notification.
5.  **Performance & Compensation:** Functions like `updateReactivationAndLastRideDates` and `compenationDueCheck` continuously update driver progress and automatically move records of eligible drivers to a "Compensation Due" sheet for payment processing.
6.  **Continuous Updates:** The system also has functions to periodically fetch and update the master data from external spreadsheets, ensuring the eligibility checks are always based on the most current information.

---

### **Technology Stack**

* **Core Language:** Google Apps Script (a JavaScript-based language for extending Google Workspace).
* **Platform:** Google Workspace (Google Sheets, Google Forms, Gmail).

---

### **Deliverables**

* Complete, well-documented Google Apps Script code.
* Instructions for setup and integration with your existing Google Sheets and Forms.
* Ongoing support and maintenance packages are available.
