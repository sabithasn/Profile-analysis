import os
from PIL import Image
import pytesseract
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
import pandas as pd

# Define keyword categories for interest inference
interest_keywords = {
    "Spirituality": ["devotion", "spiritual", "radiance", "temple", "faith"],
    "Entertainment": ["film", "movie", "tv", "viral", "LOL", "funny"],
    "Sports": ["cricket", "basketball", "team", "match", "score", "Kolkata"],
    "Creativity": ["creative", "art", "design", "frame", "candle"],
    "Academics": ["mathematics", "academy", "study", "content", "education"]
}

def infer_interests(text):
    interests = set()
    text_lower = text.lower()
    for category, keywords in interest_keywords.items():
        if any(keyword in text_lower for keyword in keywords):
            interests.add(category)
    return ", ".join(interests)

def extract_profile_data(image_folder):
    profiles = []

    for image_file in os.listdir(image_folder):
        if image_file.lower().endswith(('.png', '.jpg', '.jpeg')):
            image_path = os.path.join(image_folder, image_file)
            try:
                image = Image.open(image_path)
                text = pytesseract.image_to_string(image)

                profile = {
                    "Image File": image_file,
                    "Name": "",
                    "Bio": "",
                    "Education": "",
                    "Current City": "",
                    "Hometown": "",
                    "Relationship Status": "",
                    "Check-ins": [],
                    "Sports Teams": [],
                    "Apps and Games": [],
                    "Likes": [],
                    "Inferred Interests": infer_interests(text)
                }

                lines = text.split('\n')
                for line in lines:
                    line = line.strip()
                    if not line:
                        continue
                    if "Work at" in line:
                        profile["Bio"] = line
                    elif "Studied" in line:
                        profile["Education"] = line
                    elif "Lives in" in line:
                        profile["Current City"] = line.replace("Lives in", "").strip()
                    elif "From" in line:
                        profile["Hometown"] = line.replace("From", "").strip()
                    elif "Single" in line:
                        profile["Relationship Status"] = "Single"
                    elif any(loc in line for loc in ["MANI", "Idukki", "Yuvarani"]):
                        profile["Check-ins"].append(line)
                    elif any(team in line for team in ["Supernova", "Indian Cricket", "Kolkata"]):
                        profile["Sports Teams"].append(line)
                    elif "Basketball" in line:
                        profile["Apps and Games"].append("Basketball")
                    elif any(like in line for like in ["LOL", "Devotion", "Creative", "Candle", "Filmspace"]):
                        profile["Likes"].append(line)

                # Flatten list fields
                for key in ["Check-ins", "Sports Teams", "Apps and Games", "Likes"]:
                    profile[key] = ", ".join(profile[key])

                profiles.append(profile)

            except Exception as e:
                print(f"⚠️ Error processing {image_file}: {e}")

    return profiles

def save_to_excel(profiles, output_file):
    df = pd.DataFrame(profiles)
    df.to_excel(output_file, index=False)

# Run the script
image_folder = "."  # Current folder
output_file = "profile_analysis.xlsx"

profiles = extract_profile_data(image_folder)
save_to_excel(profiles, output_file)

print(f"✅ Excel file saved as: {output_file}")
