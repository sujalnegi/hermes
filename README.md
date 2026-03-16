<p align="center">
  <img src="static/images/android-chrome-192x192.png" alt="Hermes Logo" width="120"/>
</p>

<h1 align="center">Hermes</h1>

Hermes is a powerful, AI-driven email sorting and response generation platform. Think of it as your personal administrative assistant—designed to parse raw email files, automatically categorize them by sender and purpose, and leverage Google's Gemini AI to draft formal, context-aware replies. 

### Features

* **Email Parsing & Categorization**:
    * **Offline Processing**: Drop your raw `.eml` files into the data folder and Hermes processes them instantly.
    * **Smart Sorting**: Automatically categorizes emails by Sender Type (Professor, Student, Company, Other) and Purpose (Leave, Event, Payment, etc.).

* **Interactive Results Dashboard**:
    * **Live Statistics**: View dynamic counts of processed emails categorized beautifully.
    * **Selection Controls**: Triage your inbox easily using interactive checkboxes, Select All/Deselect All buttons, and category filters.

* **AI-Powered Response Generation**:
    * **The AI Professor Persona**: Uses `gemini-2.5-flash-lite` to draft highly formal, professional, and academic replies.
    * **Custom Rule Ingestion**: Upload custom `.pdf` guidelines (like event rules or leave policies). The AI reads these rules and strictly factors them into its generated responses natively.

* **Excel Export**:
    * **Data Structuring**: Exports all sorted emails directly into a beautifully formatted `.xlsx` workbook.
    * **Automated Column**: Responses drafted by Gemini are injected directly alongside the email metadata in a dedicated AI Generated Response column.

### Technologies Used

* **Backend**: Python, Flask
* **Frontend**: HTML5, Vanilla JavaScript, CSS3 (Glassmorphism & Modern UI)
* **AI Integration**: Google GenAI SDK (`gemini-2.5-flash-lite`)
* **Document Parsing**: `email.parser` (EML), `PyPDF2` (PDF)
* **Spreadsheet Generation**: `openpyxl`

---

## Local Setup and Installation

Follow these steps to get Hermes running on your local machine.

### 1. Prerequisites

* Python 3.7+
* pip (Python package installer)
* A valid Google Gemini API Key

### 2. Clone the Repository

Clone this repository to your local machine:

```bash
git clone https://github.com/sujalnegi/hermes.git
cd hermes
```

### 3. Install Dependencies

Install the required Python packages:

```bash
pip install -r requirements.txt
```

### 4. Configure Environment Variables

Create a `.env` file in the root directory and add your Gemini API Key:

```env
GEMINI_API_KEY=your_gemini_api_key_here
```

### 5. Run the Application

Start the Flask server:

```bash
python app.py
```

Now open your web browser and go to the following address:

`http://127.0.0.1:5000/`

You should see Hermes running!

---

## How to Use

1. **Start**: On the landing page, either drag and drop your `.eml` files into the upload zone or click **"Process Data Folder"** to parse emails currently sitting in the `data/` directory.
2. **Review & Filter**: You will be redirected to the Results page. Here you can read email previews, see their categorized tags, and filter by sender type.
3. **Select for AI**: Check the boxes next to the emails you want the AI to reply to (or use "Select All").
4. **Add Guidelines (Optional)**: Click **"Add Guidelines"** to attach a `.pdf` file containing specific rules or instructions the AI must follow when drafting replies.
5. **Export**: Click **"Export as Excel"**. The app will securely hit the Gemini API, draft the responses, and download a styled Excel report to your device.

---

## Author

* **Email**: sujal1negi@gmail.com
* **Instagram**: @sujal1negi

## Acknowledgments/Credits

* [Google Gemini AI](https://deepmind.google/technologies/gemini/) for the incredible `flash-lite` inference model.
* [Flask](https://flask.palletsprojects.com/) community.
* Vector assets inspired by mythological themes and modern web aesthetics.
