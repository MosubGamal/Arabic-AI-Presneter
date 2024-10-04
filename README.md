## Arabic AI Presenter

[![Open In Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/drive/1be6b-z-M5TRQn9ZvQg1g6qOXNVSIA2ec?usp=sharing)

# Video Demo
Watch the video of a converted PowerPoint presentation where the Arabic AI Presenter presents the slides:
<div align="center">
  <a href="https://drive.google.com/file/d/1fcSd_GyNlNJtz5qqAEQXTvnZfVQGwxRR/view?usp=sharing">
    <img src="https://img.youtube.com/vi/33DgQE6HVGY/0.jpg" alt="Watch the video">
  </a>
</div>

## Project Description
Arabic AI Presenter is a Python project that converts PowerPoint presentations into video presentations with Arabic narration. The project processes each slide of a PowerPoint presentation, extracts text descriptions, generates scripts, translates them into Arabic, and creates audio narrations. Finally, it combines the slides and audio into a video presentation.

## Installation Instructions
To run this project, you need to have Python installed on your system. Additionally, you need to install the required libraries listed below.


### How to Install the Libraries
You can install the required libraries using `pip`. Run the following command to install all the necessary libraries:

```bash
pip install colorama tqdm python-dotenv pypiwin32 edge-tts groq moviepy ipython requests
```

### Installing FFmpeg
MoviePy requires FFmpeg to be installed on your system. Follow the instructions below to install FFmpeg:

#### Windows
1. Download the FFmpeg zip file from the [official website](https://ffmpeg.org/download.html).
2. Extract the zip file to a folder (e.g., `C:\ffmpeg`).
3. Add the `bin` folder to your system's PATH environment variable:
   - Open the Start Menu and search for "Environment Variables".
   - Click on "Edit the system environment variables".
   - In the System Properties window, click on the "Environment Variables" button.
   - In the Environment Variables window, find the "Path" variable in the "System variables" section and click "Edit".
   - Click "New" and add the path to the `bin` folder (e.g., `C:\ffmpeg\bin`).
   - Click "OK" to close all windows.

#### macOS
1. Install FFmpeg using Homebrew:
   ```bash
   brew install ffmpeg
   ```

#### Linux
1. Install FFmpeg using your package manager. For example, on Ubuntu:
   ```bash
   sudo apt update
   sudo apt install ffmpeg
   ```

## Usage Instructions
1. Place your PowerPoint presentation in the project directory.
2. Update the `presentation_path` variable in `AAP.py` to the path of your PowerPoint file.
3. Run the `AAP.py` script to generate the video presentation.

```bash
python AAP.py
```

## Project Structure
- `AAP.py`: Main script that orchestrates the conversion process.
- `helpers.py`: Contains helper functions for various tasks such as converting PowerPoint to images, generating scripts, translating text, generating audio, and creating videos.
- `Slides/`: Directory where slide images and audio files are stored.
- `.env`: Environment file containing API keys.

## Example
An example PowerPoint file (`example.pptx`) is provided in the project directory. You can use this file to test the project.

## Notes
- Ensure you have the necessary API keys set in the `.env` file.
- The project uses the [Groq API](https://console.groq.com/) and [Nivdia Build](https://build.nvidia.com/) API for generating scripts and translations. You need to sign up for an API key from Groq and Nivdia Build respectively, and update the `.env` file with your API keys. Add them under the names GROQ_API_KEY, NIV_API_KEY.
