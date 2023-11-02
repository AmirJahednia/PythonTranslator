# Translator Bot

**Translator Bot** is a Python script designed to assist users in translating frequently occurring sentences within a document. It is essential to note that this tool is not an AI translator. Instead, it operates based on a manually created data source, allowing users to replace specific sentences or phrases with their corresponding translations.

# Features

- **Data Source**: Users provide a manually created data source that includes pairs of sentences in their preferred languages (e.g., English to Farsi).

- **Non-AI Translation**: Translator Bot doesn't use artificial intelligence for translation. It simply scans the document for matches and replaces the sentences based on the provided data source.

- **Customizable Language Pairs**: Translator Bot is set by default for English to Farsi translation. Users can customize it for their choice of languages by adjusting the column names in the code.

- **Case-Insensitive Matching**: Translator Bot performs case-insensitive matching, ensuring that translations are found regardless of letter case.

# Usage

1. **Set up your Data Source**: Create an Excel file with two columns, one for sentences in your preferred source language and the other for their translations. Save it in as 'DataSource.xlsx' in the script's directory.

2. **Run the Script**: Execute the script, providing the path to the document you want to translate.

3. **Review the Output**: The translated document is saved with the prefix 'Translated_' in the same directory as the original document.

# Dependencies

- Python
- pandas
- docx
  
# Project Status: Please note that this project is currently unfinished and may receive gradual updates. Feel free to submit pull requests and add your insights to the project.

# Example Data Source

| English Sentence                        | Farsi Translation                           |
|---------------------------------------|--------------------------------------------|
| This is a test.                        | این یک آزمایش است.                        |
| Please provide your name.             | لطفاً نام خود را وارد کنید.             |

## Important Note

**Translator Bot is not an AI-based translator**. It does not perform real-time translation but replaces sentences based on the provided data source. It is a handy tool for users who have a predefined set of sentences that need consistent translation within their documents.

**Customization**: By adjusting the column names in the code, users can customize Translator Bot for their choice of languages.

Feel free to contribute to this project, suggest improvements, or report issues.

# License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

P.S: The project is unfinished and might get gradual updates. feel free to submit pull requests and add your insights to the project.
