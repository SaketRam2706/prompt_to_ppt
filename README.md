# 🎯 Prompt to PPT: AI Presentation Generator

**Prompt to PPT** is a Streamlit-based web application that generates professional PowerPoint presentations using AI (via DeepSeek) and images sourced dynamically from Unsplash. With just a simple prompt, this tool creates a complete `.pptx` and optional `.pdf` presentation—styled, content-rich, and ready to download.

---

## 🚀 Features

- ✨ **AI-generated slides** from a single-line prompt  
- 🖼️ **Auto-image sourcing** via Unsplash  
- 📄 Multiple slide layouts supported  
- 🔤 Choose your **preferred font**  
- 📥 Download as `.pptx` or `.pdf`  
- 🎨 Sleek UI with custom CSS  

---

## 🛠️ Tech Stack

- **Frontend:** Streamlit  
- **AI Model:** DeepSeek Chat API  
- **Image Source:** Unsplash API  
- **Presentation Engine:** `python-pptx`  
- **PDF Conversion:** COM Automation (MS PowerPoint on Windows)

---

## 📦 Installation

### 1. Clone the Repository

```bash
git clone https://github.com/yourusername/prompt_to_ppt.git
cd prompt_to_ppt
```
### 2. Set Up Environment Variables
Create a .env file in the root directory with the following:
```
DEEPSEEK_API_KEY=your_deepseek_api_key
UNSPLASH_ACCESS_KEY=your_unsplash_access_key
```


### 3. Install Dependencies

```
pip install -r requirements.txt

```

### 4. Running the App

```
streamlit run app.py

```

## 📤 Output
The generated .pptx will be available for download and can be saved locally.

## Requirements
- Python 3.7+

- Streamlit

- Access to DeepSeek & Unsplash API keys

## Acknowledgements
- DeepSeek AI

- Unsplash API

- Streamlit
