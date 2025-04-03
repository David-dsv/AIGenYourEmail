# AIGenYourEmail

## Description
AIGenYourEmail is an automation tool based on the Azure OpenAI API that generates personalized professional emails for any type of client. Using a GPT-4o model, this tool dynamically adapts an email template based on specific recipient information.

## Features
- **Read clients from an Excel file**: Extract client information from a `clients.xlsx` file.
- **Automatically generate personalized emails**: Use GPT-4o to tailor the template according to the client's sector, country, and needs.
- **Save generated emails**: Store personalized emails as `.txt` files.
- **Export client data in Excel format**: Extract and structure data from a `.txt` file to `.xlsx` with advanced formatting.

## Technologies Used
- **Python**
- **Azure OpenAI API (GPT-4o)**
- **Pandas** for Excel file manipulation
- **Requests** for API interaction
- **Dotenv** for API key management
- **Openpyxl** for advanced Excel formatting

## Installation
1. **Clone the repository**
   ```sh
   git clone https://github.com/David-dsv/AIGenYourEmail.git
   cd AIGenYourEmail
   ```
2. **Create and activate a virtual environment**
   ```sh
   python3 -m venv env
   source env/bin/activate  # For macOS/Linux
   env\Scripts\activate    # For Windows
   ```
3. **Install dependencies**
   ```sh
   pip install -r requirements.txt
   ```
4. **Create a `.env` file and add your Azure API keys**
   ```ini
   AZURE_OPENAI_KEY=YourAPIKey
   AZURE_OPENAI_ENDPOINT=YourAzureEndpoint
   ```

## Usage
1. **Run the main script**
   ```sh
   python main.py
   ```
2. **Generated emails will be saved in the `mails_personnalises/` folder**
3. **Client information will be exported to `clients.xlsx`**

## Example of Generated Emails
A typical email might look like this:
```
Hello Mr. Smith,
I hope this message finds you well...
...
Best regards,
David Vuong
Sales Manager
```

## Contributing
1. **Fork the project**
2. **Create a feature branch**
   ```sh
   git checkout -b my-new-feature
   ```
3. **Submit a pull request**

## Authors
- **David Vuong** - [GitHub](https://github.com/David-dsv)

## License
This project is licensed under the MIT License - see the `LICENSE` file for details.

## About Me
I'm **David Soeiro-Vuong**, a third-year Computer Science student working as an apprentice at **TW3 Partners**, a company specialized in **Generative AI**. Passionate about artificial intelligence and language model optimization, I focus on creating efficient model merges that balance performance and capabilities.

🔗 [Connect with me on LinkedIn](https://www.linkedin.com/in/david-soeiro-vuong-a28b582ba/)

🔗 [Follow me on Hugging Face](https://huggingface.co/Davidsv/)

