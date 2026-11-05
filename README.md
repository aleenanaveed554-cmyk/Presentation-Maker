# ✨ AI Presentation Maker

A professional AI-powered presentation generator built with Python,
Streamlit, Groq API and python-pptx.

## 🚀 Features

- AI-powered presentation generation
- Professional PowerPoint design
- Multiple presentation styles
- Multiple audience types
- Multiple tones
- Multiple languages
- Adjustable slide count
- AI-generated slide content
- Speaker notes
- Real PowerPoint (.pptx) generation
- Download presentation
- Streamlit Cloud compatible
- Secure API key management

## 🛠️ Technologies

- Python
- Streamlit
- Groq API
- python-pptx

## 📁 Project Structure

AI-Presentation-Maker/

├── app.py
├── requirements.txt
├── README.md
├── .gitignore
└── .streamlit/
    └── secrets.toml.example

## 🔑 API Key Setup

Create a Groq API key.

For local development create:

.streamlit/secrets.toml

Add:

GROQ_API_KEY = "your_groq_api_key_here"

Never upload your real API key to GitHub.

## ▶️ Run Locally

Install dependencies:

pip install -r requirements.txt

Run:

streamlit run app.py

## ☁️ Streamlit Cloud Deployment

1. Push the project to GitHub.
2. Open Streamlit Community Cloud.
3. Create a new app.
4. Select your GitHub repository.
5. Select app.py.
6. Deploy the application.
7. Open the app's Settings / Secrets section.
8. Add:

GROQ_API_KEY = "your_groq_api_key_here"

9. Save the secret.
10. Reboot/redeploy the application.

## 🎯 Example Topics

Artificial Intelligence in Education

Digital Marketing Strategy

Climate Change

Cybersecurity

Future of Technology

Machine Learning

Business Management

Social Media Marketing
