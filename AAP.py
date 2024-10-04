import os
import asyncio
import shutil
from helpers import ppt_to_png, slide_descriptions, slide_scripts, slide_translate, slide_audio, create_and_play_video, play_video
from helpers import GroqClient
from colorama import Fore, Style, init
from tqdm import tqdm

# Initialize colorama
init(autoreset=True)

def clear_slides_folder(folder_path):
    """
    Function to delete all files in the specified folder.
    """
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        try:
            if os.path.isfile(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            print(f'Failed to delete {file_path}. Reason: {e}')

async def main(presentation_path):
    # Specify the full path to the PowerPoint presentation
    
    slides_folder = "Slides"
    # Check if the folder exists, and if not, create it
    if not os.path.exists(slides_folder):
        os.makedirs(slides_folder)
    
    # Clear the Slides folder before starting the process
    print(Fore.BLUE + "Clearing the slides folder..." + Style.RESET_ALL)
    clear_slides_folder(slides_folder)
    print(Fore.GREEN + "Slides folder cleared." + Style.RESET_ALL)
    
    # Convert the PowerPoint presentation to PNG images, one for each slide.
    print(Fore.BLUE + "Converting PowerPoint to PNG images..." + Style.RESET_ALL)
    ppt_to_png(presentation_path)

    print(Fore.GREEN + "Conversion completed." + Style.RESET_ALL)
    
    # Extract text descriptions from each slide in the PowerPoint presentation.
    print(Fore.BLUE + "Extracting text descriptions from slides..." + Style.RESET_ALL)
    slides_dict = slide_descriptions(slides_folder)
    print(Fore.GREEN + "Text descriptions extracted." + Style.RESET_ALL)
    
    # System prompt for the AI script writer; instructs the AI to create a smooth script based on slide descriptions.
    systm_prompt = "You are an AI presenter. You will be given a text description of each slide. Create a smooth script to descript each slide. Only output the script not the instruction how to present."
    
    # Instantiate the AI script writer client with the provided system prompt.
    script_writer = GroqClient(systm_prompt)
    
    # Generate scripts for each slide based on the descriptions.
    print(Fore.BLUE + "Generating scripts for slides..." + Style.RESET_ALL)
    slides_dict = slide_scripts(slides_dict)
    print(Fore.GREEN + "Scripts generated." + Style.RESET_ALL)
    
    # Translate scripts to the desired language.
    print(Fore.BLUE + "Translating scripts to the desired language..." + Style.RESET_ALL)
    slides_dict = slide_translate(slides_dict)
    print(Fore.GREEN + "Translation completed." + Style.RESET_ALL)
    
    # Generate audio for each slide's script asynchronously.
    print(Fore.BLUE + "Generating audio for each slide's script..." + Style.RESET_ALL)
    slides_dict = await slide_audio(slides_dict)
    print(Fore.GREEN + "Audio generated." + Style.RESET_ALL)
    
    # Create a video from the slides and corresponding audio, then get the path to the created video.
    print(Fore.BLUE + "Creating video from slides and audio..." + Style.RESET_ALL)
    print(slides_dict)
    output_path = create_and_play_video(slides_dict)
    print(Fore.GREEN + "Video created." + Style.RESET_ALL)

    # Play the generated video.
    # print(Fore.BLUE + "Playing the video..." + Style.RESET_ALL)
    # play_video(output_path)
    # print(Fore.GREEN + "Video playback started." + Style.RESET_ALL)

# Run the main async function
presentation_path = input("Enter the full path to the PowerPoint presentation: ")
asyncio.run(main(presentation_path))
