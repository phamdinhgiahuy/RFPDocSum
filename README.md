<div align="left" style="position: relative;">
<img src="https://raw.githubusercontent.com/PKief/vscode-material-icon-theme/ec559a9f6bfd399b82bb44393651661b08aaf7ba/icons/folder-markdown-open.svg" align="right" width="30%" style="margin: -20px 0 0 20px;">
<h1>RFPDOCSUM</h1>
<p align="left">
	<em>Automate the consolidation and analysis of Request For Proposal (RFP) documents</em>
</p>
<p align="left">
	<img src="https://img.shields.io/github/license/phamdinhgiahuy/RFPDocSum?style=flat&logo=opensourceinitiative&logoColor=white&color=0080ff" alt="license">
	<img src="https://img.shields.io/github/last-commit/phamdinhgiahuy/RFPDocSum?style=flat&logo=git&logoColor=white&color=0080ff" alt="last-commit">
	<img src="https://img.shields.io/github/languages/top/phamdinhgiahuy/RFPDocSum?style=flat&color=0080ff" alt="repo-top-language">
	<img src="https://img.shields.io/github/languages/count/phamdinhgiahuy/RFPDocSum?style=flat&color=0080ff" alt="repo-language-count">
</p>
<p align="left">Built with the tools and technologies:</p>
<p align="left">
	<img src="https://img.shields.io/badge/Streamlit-FF4B4B.svg?style=flat&logo=Streamlit&logoColor=white" alt="Streamlit">
	<img src="https://img.shields.io/badge/Python-3776AB.svg?style=flat&logo=Python&logoColor=white" alt="Python">
	<img src="https://img.shields.io/badge/pandas-150458.svg?style=flat&logo=pandas&logoColor=white" alt="pandas">
</p>
</div>
<br clear="right">

## 🔗 Table of Contents

- [🔗 Table of Contents](#-table-of-contents)
- [📍 Overview](#-overview)
  - [Project Objectives](#project-objectives)
    - [Key Functionalities:](#key-functionalities)
    - [Benefits:](#benefits)
- [👾 Features](#-features)
  - [⚙️ RFP event configuration](#️-rfp-event-configuration)
  - [💲 Pricing](#-pricing)
  - [❔ Questionnaire](#-questionnaire)
- [🎥 Demo](#-demo)
- [📁 Project Structure](#-project-structure)
  - [📂 Project Index](#-project-index)
- [🚀 Getting Started](#-getting-started)
  - [☑️ Prerequisites](#️-prerequisites)
  - [⚙️ Installation](#️-installation)
  - [🤖 Usage](#-usage)
- [📌 Project Roadmap](#-project-roadmap)
- [🔰 Contributing](#-contributing)
- [🎗 License](#-license)
- [🙌 Acknowledgments](#-acknowledgments)

---

## 📍 Overview

### Project Objectives  

This project aims to **reduce manual effort** and **enhance efficiency** by developing an automated document summarization tool.  

#### Key Functionalities:  
- **Automated data aggregation**: Compiling vendor responses from questionnaires and pricing forms.  
- **Summarization**: Generating concise, actionable summaries to support strategic initiatives.  
- **High-level insights**: Providing summaries tailored for decision-makers, enabling quicker, data-driven decisions.  

#### Benefits:  
By streamlining the RFP response aggregation process, this tool will:  
- Improve scalability.
- Free up valuable time and resources.  
- Enhance decision-making for strategic planning.  

This project represents a significant step forward in optimizing procurement workflows and empowering organizations to focus on strategic objectives rather than manual administrative tasks.  

---

## 👾 Features

### ⚙️ RFP event configuration
- **Event Options**:
  - **Single File:** Prices and questionnaires are combined into a single response file.
  - **Separate Files:** Prices and questionnaires are stored in separate files.
- **Upload Files:** Accepts valid `.xlsx` files for easy file and supplier configuration.
- **Event and Supplier Names:** Displays in the final consolidated file.

### 💲 Pricing
- **Organized Data**: Differentiates descriptions and prices.  
- **Color-Coded Pricing**: Highlights vendor-specific values.  
- **Aggregation Options**:  
  - **Side-by-Side**: Prices from multiple vendors in one sheet.  
  - **Sheet-by-Sheet**: Vendor prices in separate sheets.  
- **Analysis Tools**: Generates summaries and diagrams.

### ❔ Questionnaire
- **Parse Responses**: Matches template columns, highlights mismatched rows, and extracts vendor data.  
- **Consolidation Options**:  
  - **Side-by-Side**: All responses in one sheet.  
  - **Separate Sheets**: Each vendor's data in its own sheet.  
- **Summarization**: Option to create concise summaries.

---

## 🎥 Demo

Below is a demonstration of how the tool works:

![Demo of RFP Tool](demo/demoRFP.gif)

---

## 📁 Project Structure

```sh
└── RFPDocSum/
    ├── LICENSE
    ├── README.md
    ├── main_app.py
    ├── requirements.txt
    └── tools
        ├── consolidate.py
        └── event_config.py
```


### 📂 Project Index
<details open>
	<summary><b><code>RFPDOCSUM/</code></b></summary>
	<details> <!-- __root__ Submodule -->
		<summary><b>__root__</b></summary>
		<blockquote>
			<table>
			<tr>
				<td><b><a href='https://github.com/phamdinhgiahuy/RFPDocSum/blob/master/main_app.py'>main_app.py</a></b></td>
				<td>Run this file with streamlit to start the app</td>
			</tr>
			<tr>
				<td><b><a href='https://github.com/phamdinhgiahuy/RFPDocSum/blob/master/requirements.txt'>requirements.txt</a></b></td>
				<td>install dependencies with pip</td>
			</tr>
			</table>
		</blockquote>
	</details>
</details>

---
## 🚀 Getting Started

### ☑️ Prerequisites

Before getting started with RFPDocSum, ensure your runtime environment meets the following requirements:

- **Programming Language:** Python
- **Package Manager:** Pip


### ⚙️ Installation

Install RFPDocSum using one of the following methods:

**Build from source:**

1. Clone the RFPDocSum repository:
```sh
❯ git clone https://github.com/phamdinhgiahuy/RFPDocSum
```

2. Navigate to the project directory:
```sh
❯ cd RFPDocSum
```

3. Install the project dependencies:


**Using `pip`** &nbsp; [<img align="center" src="https://img.shields.io/badge/Pip-3776AB.svg?style={badge_style}&logo=pypi&logoColor=white" />](https://pypi.org/project/pip/)

```sh
❯ pip install -r requirements.txt
```




### 🤖 Usage
Run RFPDocSum using the following command:
**Using `pip`** &nbsp; [<img align="center" src="https://img.shields.io/badge/Pip-3776AB.svg?style={badge_style}&logo=pypi&logoColor=white" />](https://pypi.org/project/pip/)

```sh
❯ streamlit run main_app.py
```

---
## 📌 Project Roadmap

- [X] **`Task 1`**: <strike>Implement Consolidation features</strike>
- [X] **`Task 2`**: <strike>Implement Pricing Analysis and Summary features</strike>
- [X] - [ ] **`Task 3`**: <strike>Implement Questionnaire summary and supplier modification of template features</strike>
- [ ] **`Task 4`**: IReplace pure NLP models with LLM for ehanced summarization and interaction with the consolidated content.

---

## 🔰 Contributing

- **💬 [Join the Discussions](https://github.com/phamdinhgiahuy/RFPDocSum/discussions)**: Share your insights, provide feedback, or ask questions.
- **🐛 [Report Issues](https://github.com/phamdinhgiahuy/RFPDocSum/issues)**: Submit bugs found or log feature requests for the `RFPDocSum` project.
- **💡 [Submit Pull Requests](https://github.com/phamdinhgiahuy/RFPDocSum/blob/main/CONTRIBUTING.md)**: Review open PRs, and submit your own PRs.

<details closed>
<summary>Contributing Guidelines</summary>

1. **Fork the Repository**: Start by forking the project repository to your github account.
2. **Clone Locally**: Clone the forked repository to your local machine using a git client.
   ```sh
   git clone https://github.com/phamdinhgiahuy/RFPDocSum
   ```
3. **Create a New Branch**: Always work on a new branch, giving it a descriptive name.
   ```sh
   git checkout -b new-feature-x
   ```
4. **Make Your Changes**: Develop and test your changes locally.
5. **Commit Your Changes**: Commit with a clear message describing your updates.
   ```sh
   git commit -m 'Implemented new feature x.'
   ```
6. **Push to github**: Push the changes to your forked repository.
   ```sh
   git push origin new-feature-x
   ```
7. **Submit a Pull Request**: Create a PR against the original project repository. Clearly describe the changes and their motivations.
8. **Review**: Once your PR is reviewed and approved, it will be merged into the main branch. Congratulations on your contribution!
</details>

<details closed>
<summary>Contributor Graph</summary>
<br>
<p align="left">
   <a href="https://github.com{/phamdinhgiahuy/RFPDocSum/}graphs/contributors">
      <img src="https://contrib.rocks/image?repo=phamdinhgiahuy/RFPDocSum">
   </a>
</p>
</details>

---

## 🎗 License

This project is protected under the [MIT License](https://choosealicense.com/licenses/mit/) License. For more details, refer to the [LICENSE](https://github.com/phamdinhgiahuy/RFPDocSum/blob/main/LICENSE) file.

---

## 🙌 Acknowledgments


---