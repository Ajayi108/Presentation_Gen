# AI Presentation Generator

A Streamlit app that turns a single topic into a PowerPoint presentation using Gemini for slide content and Unsplash for slide images.

## What it does

- Loads API keys from `.env`
- Sends a prompt to Gemini for structured slide JSON
- Optionally fetches matching images from Unsplash for each slide
- Previews the generated slide outline and images in Streamlit
- Creates a `.pptx` file in `generated_presentations/`
- Lets you download the same `.pptx` file from the browser

## Setup

1. Create and activate a virtual environment.
2. Install dependencies:
   `pip install -r requirements.txt`
3. Add your API keys to `.env`:
   `GEMINI_API_KEY=your_key_here`
   `UNSPLASH_ACCESS_KEY=your_unsplash_key_here`

You can also use `GOOGLE_API_KEY` instead of `GEMINI_API_KEY`.

## Run

```bash
streamlit run app.py
```

## Output

After you click `Generate presentation`, the app will:

- save the PowerPoint file in `generated_presentations/`
- show the saved file path in the UI
- provide a download button for the same file
- embed Unsplash images when that option is enabled

## Notes

- The app defaults to the `gemini-2.5-flash` model.
- Export requires `python-pptx` to be installed.
- The API keys are never shown in the UI.
- The code includes inline comments around the main generation, image, and export steps. 
