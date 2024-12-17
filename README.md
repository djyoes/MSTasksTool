# Microsoft Post Migration Helper Wizard (MSTasksTool)

[![GitHub stars](https://img.shields.io/github/stars/djyoes/MSTasksTool?style=social)](https://github.com/djyoes/MSTasksTool/stargazers)
[![GitHub issues](https://img.shields.io/github/issues/djyoes/MSTasksTool)](https://github.com/djyoes/MSTasksTool/issues)
[![GitHub license](https://img.shields.io/github/license/djyoes/MSTasksTool)](https://github.com/djyoes/MSTasksTool/blob/main/LICENSE)
[![GitHub last commit](https://img.shields.io/github/last-commit/djyoes/MSTasksTool)](https://github.com/djyoes/MSTasksTool/commits/main)

The **Microsoft Post Migration Helper Wizard (MSTasksTool)** is a robust, user-friendly application designed to streamline the process of configuring and transitioning end-user environments to a Microsoft-centric ecosystem. It provides both IT professionals and end-users with a clear, intuitive interface to complete necessary setup tasks with ease. This tool also supports customization, allowing IT teams to embed company-specific documentation, guides, and communication plans directly into the package.

---

## 🚀 Highlights

*   **User-Friendly Interface:** The application provides clear, step-by-step guidance, designed for both IT professionals and end-users.
*   **Streamlined Actions:**  Perform common, time-consuming tasks with simple, single-click buttons.
*   **Comprehensive Logging:** Track all actions with detailed logs for effective troubleshooting and auditing.
*   **Customizable Deployment:** Integrate company-specific documentation, support guides, and communication plans for seamless user support.
*   **Robust validation**: The application is designed to validate every step, and provide you with correct and helpful messages at every stage of the process.
*  **Enhanced Edge Experience**: The edge reset functionality has been improved and now automatically relaunches edge after the profile is reset.

---

## ✨ Features

*   ✅ **1. Add Work/School Account:** Seamlessly integrate Microsoft work or school accounts. The application will check if an account is already set, and if it is then it will display the current tenant name, otherwise the application will open the settings to add a work or school account.
*   ✅ **2. Clear Outlook Profiles:** Remove existing Outlook profiles to start fresh. The application will not attempt to remove the profile, if no profiles are found.
*   ✅ **3. Reset Edge:**  Forcefully closes all instances of Edge, clears user data, and automatically relaunches Edge, it will now also log if an error occurs.
*   ✅ **4. Set Edge as Default:** Opens the Default Apps settings for you to easily configure Microsoft Edge as the default browser, and validates that Edge is set as default.
*   ✅ **5. Setup OneDrive:** Sets up OneDrive for cloud storage and file synchronization. This now uses the registry to validate if one drive has been configured, and also will configure known folder move.
*   ✅ **Logs & Troubleshooting:**  Detailed logs for every step, allowing for troubleshooting.
*   ✅ **Customizable Packaging:** Add your IT-specific documentation, and guides for seamless end-user support.
*   ✅ **Save Functionality:** The save function now validates that all tasks are complete and will then exit the application and will provide an exit code of `0` indicating success to the deployment tool, if all tests are done correctly, and will save a log of everything that has happened in the application.

---

## 🖼 Screenshots

![image](https://github.com/user-attachments/assets/bfeee219-7819-40f6-8bf8-6ab05b4d3505)

---

## 🛠 Prerequisites

*   **Operating System:** Windows 10 or Windows 11 (both 32-bit and 64-bit versions)
*   **Python:** Version 3.8 or higher *(if running from source code)*

---

## 🔧 Getting Started

### 🎯 Executable Distribution

A standalone executable version is available, so that users can easily use the tool without having to install python.

#### Steps to Use the Executable:

1.  Download the **MSTasksTool.exe** from the [official releases page](https://github.com/djyoes/MSTasksTool/releases).
2.  Double-click the executable to launch the wizard.
3.  Follow the on-screen instructions to perform migration tasks.
4.  Optionally save the application log for analysis and troubleshooting, and also view any included company-specific documentation or guides that have been packaged with the application.

---

### 🐍 Running from Source Code

To run from the source code, or to contribute to the development you can use the following steps.

#### 1. Clone the Repository:

```bash
git clone https://github.com/djyoes/MSTasksTool.git
```

#### 2. Navigate to the Project Directory:

```bash
cd MSTasksTool
```

#### 3. Install Dependencies:

```bash
pip install -r requirements.txt
```

#### 4. Run the Tool:

```bash
python MSTasksTool.py
```

#### 5. Follow On-Screen Instructions:

*   Use the intuitive interface to perform the required tasks, and check the progress of each task, before saving.
*  You can then view the logs and save them to a local file for analysis.

---

## 💻 Usage

1.  **Launch the Tool:**
    *   Start the tool from the executable or source code.
2.  **Run Tasks:**
    *   Use the numbered buttons to perform actions like adding work accounts, resetting Edge, or setting up OneDrive.
3.  **View Progress:**
    *   Monitor the real-time status of each task.
4. **Access Custom Documentation:**
    * IT admins can package guides or communication plans for user reference
5. **Access Logs:**
    *  View detailed logs by clicking the "View Logs" button.
    *   Use the "Save Logs (prompt)" button to save the logs for troubleshooting purposes.
6.  **Complete Your Tasks:**
    *  Validate that all the tasks have been run correctly, by checking the task status, and when all steps are complete, then you can save the progress and exit the tool.

---

## ⭐ Contributing

Contributions are welcome! If you want to contribute to this project, then please use the following steps:

1. Fork the repository.
2. Create a feature branch:

```bash
git checkout -b feature-name
```

3. Make your changes, and commit them.
4. Submit a pull request.

For feature requests or bug reports, open an issue [here](https://github.com/djyoes/MSTasksTool/issues).

---

## 📄 License

This project is licensed under the **MIT License**. See the [LICENSE](LICENSE) file for details.

---

## 📢 Contact

*   **GitHub**: [MSTasksTool](https://github.com/djyoes/MSTasksTool)
*   **Email**: [djyoes@gmail.com](mailto:djyoes@gmail.com)

---

### ⭐ Star the Repo

If you find the **Microsoft Post Migration Helper Wizard** helpful, please give it a ⭐ to support the project, and help spread the word!

---

## 🔍 Keywords

`#windows-tools` `#microsoft-setup` `#sysadmin-tools` `#onedrive-setup` `#edge-browser` `#python-automation` `#it-tools`
```

**Key Improvements in this README:**

*   **More Descriptive:** The text now provides more information about each feature and task, and also the intended functionality of the save button.
*   **Comprehensive:** The steps on how to run it from source, and also from the executable are now more detailed.
*  **Robustness Highlighted**: The added information about the robustness of the tool, and the error handling should provide users with a more reliable experience.
*   **Better UI Feedback**:  The added information about how the UI is updated should provide users with confidence that the tool is performing actions as expected.
*   **Clear Flow:** The usage section has been updated to show users the correct flow for using the application.
*   **More accurate:** The information provided should now reflect the current state of the application.
*  **Actionable Language:** The language has been updated to make it more appealing and effective for end users and IT professionals.

This new README.md should be more professional, comprehensive, and user friendly. It is also more complete and accurately reflects the current functionality of the `MSTasksTool`.

Please let me know if there is anything else you need!
