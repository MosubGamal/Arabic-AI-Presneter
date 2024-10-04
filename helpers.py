import win32com.client
import os
from dotenv import load_dotenv
import base64
import edge_tts
from groq import Groq
from moviepy.editor import ImageClip, AudioFileClip, concatenate_videoclips
from IPython.display import Video
import requests
import zipfile
import json

# Load .env file
load_dotenv()

# Get the API key
api_key = os.getenv('GROQ_API_KEY')
niv_api_key = os.getenv('NIV_API_KEY')

def ppt_to_png(presentation_path):
    try:
        Application = win32com.client.Dispatch("PowerPoint.Application")
        Presentation = Application.Presentations.Open(presentation_path, WithWindow=False)
        slides_folder = os.path.join(os.path.dirname(presentation_path), "Slides")
        if not os.path.exists(slides_folder):
            os.makedirs(slides_folder)
        for i, slide in enumerate(Presentation.Slides):
            image_path = os.path.join(slides_folder, f"{i + 1}.jpg")
            slide.Export(image_path, "JPG")
        Presentation.Close()
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        Application.Quit()


def encode_image(image_path):
    with open(image_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode('utf-8')


def get_image_chat_content(image_path):
    base64_image = encode_image(image_path)
    client = Groq(api_key=api_key)
    chat_completion = client.chat.completions.create(
        messages=[
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": "Get the text from the image"},
                    {
                        "type": "image_url",
                        "image_url": {
                            "url": f"data:image/jpeg;base64,{base64_image}",
                        },
                    },
                ],
            }
        ],
        model="llava-v1.5-7b-4096-preview",
        temperature=0,
        max_tokens=256,
    )
    return chat_completion.choices[0].message.content

def get_image_chat_content_phi(image_path, niv_api_key, invoke_url="https://ai.api.nvidia.com/v1/vlm/microsoft/phi-3-vision-128k-instruct", stream=False):
    # Read image from local path
    """
    Uploads an image to NVIDIA's VLM service and processes it using the Florence model.

    Args:
        image_path (str): Path to the image file to be processed.
        niv_api_key (str): NVIDIA API key.
        invoke_url (str, optional): URL to invoke the VLM service. Defaults to "https://ai.api.nvidia.com/v1/vlm/microsoft/phi-3-vision-128k-instruct".
        stream (bool, optional): Whether to stream the output. Defaults to False.

    Returns:
        str: The extracted text from the image.
    """
    with open(image_path, "rb") as f:
        image_b64 = base64.b64encode(f.read()).decode()

    # assert len(image_b64) < 180_000, \
    #     "To upload larger images, use the assets API (see docs)"

    headers = {
        "Authorization": f"Bearer {niv_api_key}",
        "Accept": "text/event-stream" if stream else "application/json"
    }

    payload = {
        "messages": [
            {
                "role": "user",
                "content": f'Get all text <img src="data:image/png;base64,{image_b64}" />'
            }
        ],
        "max_tokens": 512,
        "temperature": 1.00,
        "top_p": 0.70,
        "stream": stream
    }

    response = requests.post(invoke_url, headers=headers, json=payload)
 
    output = response.json()
    return output["choices"][0]["message"]["content"]

def get_image_chat_content_florence(image_path, output_path, niv_api_key):
    """
    Uploads an image to NVIDIA's VLM service and processes it using the Florence model.

    Args:
        image_path (str): The path to the image file to upload and process.
        output_path (str): The path to save the generated video (without the .zip extension).
        niv_api_key (str): The API key for accessing the NVIDIA VLM service.

    Returns:
        str: The generated video's captions.
    """
    def _upload_asset(input, description, header_auth):
        authorize = requests.post(
            "https://api.nvcf.nvidia.com/v2/nvcf/assets",
            headers={
                "Authorization": header_auth,
                "Content-Type": "application/json",
                "accept": "application/json",
            },
            json={"contentType": "image/jpeg", "description": description},
            timeout=30,
        )
        authorize.raise_for_status()

        response = requests.put(
            authorize.json()["uploadUrl"],
            data=input,
            headers={
                "x-amz-meta-nvcf-asset-description": description,
                "content-type": "image/jpeg",
            },
            timeout=300,
        )

        response.raise_for_status()
        return str(authorize.json()["assetId"])

    def _generate_content(asset_id):
        prompt = "<MORE_DETAILED_CAPTION>"
        content = f'{prompt}<img src="data:image/jpeg;asset_id,{asset_id}" />'
        return content

    nvai_url = "https://ai.api.nvidia.com/v1/vlm/microsoft/florence-2"
    header_auth = f"Bearer {niv_api_key}"
    
    with open(image_path, "rb") as image_file:
        asset_id = _upload_asset(image_file, "Test Image", header_auth)
    
    content = _generate_content(asset_id)
    
    inputs = {
        "messages": [{
            "role": "user",
            "content": content
        }]
    }

    headers = {
        "Content-Type": "application/json",
        "NVCF-INPUT-ASSET-REFERENCES": asset_id,
        "NVCF-FUNCTION-ASSET-IDS": asset_id,
        "Authorization": header_auth,
        "Accept": "application/json"
    }

    response = requests.post(nvai_url, headers=headers, json=inputs)
    
    with open(f"{output_path}.zip", "wb") as out:
        out.write(response.content)

    with zipfile.ZipFile(f"{output_path}.zip", 'r') as z:
        file_list = z.namelist()
        
        # Filter out files that end with .response
        response_files = [file for file in file_list if file.endswith('.response')]
        
        if response_files:
            # Extract and read each .response file
            for response_file in response_files:
                # Extract the specific file
                z.extract(response_file, 'Slides/florence')
                file_path = f'Slides/florence/{response_file}'
                
                # Read and process the .response file
                with open(file_path, 'rb') as f:
                    data = f.read()
                    data = data.decode('utf-8')
                    response_json = json.loads(data)

    return response_json["choices"][0]["message"]["content"][len("<MORE_DETAILED_CAPTION>"):]


def slide_descriptions(folder_path):
    slide_description = {}
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        if os.path.isfile(file_path) and file_path.lower().endswith((".png", ".jpg", ".jpeg", ".gif", ".bmp")):
            try:
                # content = get_image_chat_content(file_path) # use llava
                # content = get_image_chat_content_phi(file_path, niv_api_key) # use phi
                try:
                    content = get_image_chat_content_phi(file_path, niv_api_key)
                except:
                    output_file = "Slides/florence"
                    if not os.path.exists(output_file):
                        os.makedirs(output_file)
                    content = get_image_chat_content_florence(file_path, output_file, niv_api_key) # use florence
                key = int(filename.split('.')[0])
                slide_description[key] = {"image_path": f"Slides/{filename}", "slide_description" : content}
            # print(slide_description[key])
            except Exception as e:
                print(f"Error processing {filename}: {e}")
                slide_description[key] = "Error in processing"
    slide_description = dict(sorted(slide_description.items(), key=lambda item: item[0]))
    return slide_description


class GroqClient:
    def __init__(self, sytem_prompt, temperature = 0.5, max_token=256):
        self.api_key = api_key
        self.client = Groq(api_key=api_key)
        self.system_prompt = sytem_prompt
        self.conversation_history = []
        self.conversation_history.append({"role": "system", "content": self.system_prompt})
        self.max_token = max_token
        self.temperature = temperature

    def send_message(self, content, role="user", single_mode=False):
        if single_mode:
            self.conversation_history = []
            self.conversation_history.append({"role": "system", "content": self.system_prompt})
        self.conversation_history.append({"role": role, "content":  content})

    def get_response(self):
        chat_completion = self.client.chat.completions.create(
            messages=self.conversation_history,
            model="llama-3.1-70b-versatile",
            temperature=self.temperature,
            max_tokens=self.max_token,
        )
        response = chat_completion.choices[0].message.content
        self.conversation_history.append({"role": "assistant", "content": response})
        return response

    def get_conversation_history(self):
        return self.conversation_history


def slide_scripts(slide_description, end_slide=None):
    systm_prompt = "You are an AI presenter. You will be given a text description of each slide. Create a smooth script to descript each slide. Only output the script not the instruction how to present"
    script_writter = GroqClient(systm_prompt)
    for key, value in slide_description.items():
        script_writter.send_message("Slide Description: "+ value["slide_description"])
        response = script_writter.get_response()
        slide_description[key]["slide_script"] = response
        # print(slide_description[key])
        if key == end_slide:
            break
    return slide_description


def slide_translate(slides_dict):
    systm_prompt = "Translate the text to Arabic. do not include the table focus only on the insight from it. The text will be used for tts, so do not include any other text. "
    arabic_translator = GroqClient(systm_prompt, temperature=0, max_token=400)
    for key, value in slides_dict.items():
        slide_script = value.get("slide_script")
        if slide_script:
            arabic_translator.send_message(slide_script, single_mode=True)
            response = arabic_translator.get_response()
            slides_dict[key]["arabic_script"] = response
            # print(slides_dict[key])
    return slides_dict


async def generate_audio(text, voice_name="ar-BH-AliNeural", output_file="output_audio.mp3"):
    communicate = edge_tts.Communicate(text=text, rate="+30%", voice=voice_name)
    await communicate.save(output_file)
    print(f"Audio saved to {output_file}")
    return output_file


async def slide_audio(slides_dict):
    for key, value in slides_dict.items():
        slide_script = value.get("arabic_script")
        if slide_script:
            audio_file = await generate_audio(slide_script,  voice_name="ar-EG-SalmaNeural", output_file=f"Slides/{key}.mp3")
            slides_dict[key]["audio_path"] = audio_file
    return slides_dict


def create_and_play_video(slides, output_path="final_presentation.mp4", fps=1):
    video_clips = []
    for key, value in slides.items():
        img_clip = ImageClip(value['image_path'])
        audio_clip = AudioFileClip(value['audio_path'])
        img_clip = img_clip.set_duration(audio_clip.duration)
        video_clip = img_clip.set_audio(audio_clip)
        video_clips.append(video_clip)
    final_video = concatenate_videoclips(video_clips, method="compose")
    final_video.write_videofile(output_path, fps=fps)
    print("Video created successfully!")
    return output_path


def play_video(video_path):
    return Video(video_path, embed=True)
