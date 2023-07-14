import os
import sys
import time
import requests
import pystray
from PIL import Image, ImageDraw, ImageFont, ImageColor
import threading
import tkinter as tk
from tkinter import simpledialog
from win32com.client import Dispatch
from tkinter import colorchooser



# Global variables
nightscout_url = None
url_file = "urls.txt"
last_blood_glucose = None
last_trend_arrow = None
last_value_timestamp = None
delta = None
unit = "mg/dL"
target_range = (70, 180)
target_colors = ('red', 'yellow')


def read_urls_from_file():
    urls = []

    if os.path.isfile(url_file):
        with open(url_file, "r") as file:
            urls = file.readlines()

    return [url.strip() for url in urls if url.strip()]


def save_url_to_file(url):
    with open(url_file, "a") as file:
        file.write(url + "\n")

def get_text_color(blood_glucose):
    global target_range, target_colors

    # Convert the range values to integers
    lower_range = int(target_range[0])
    upper_range = int(target_range[1])

    lower_color, upper_color = target_colors

    if blood_glucose < lower_range:
        return lower_color
    elif blood_glucose > upper_range:
        return upper_color
    else:
        return "white"


def create_icon_image(blood_glucose, trend_arrow):

    icon_size = (64, 64)
    image = Image.new("RGBA", icon_size, (255, 255, 255, 0))
    draw = ImageDraw.Draw(image)

    font_size = 37
    font = ImageFont.truetype("arial.ttf", font_size)
    text_color = get_text_color(blood_glucose)

    text_width, text_height = draw.textbbox((0, 0), str(blood_glucose), font=font)[2:]
    text_x = (icon_size[0] - text_width) // 2
    text_y = 0

    draw.text((text_x, text_y), str(blood_glucose), font=font, fill=text_color, stroke_width=1, stroke_fill=text_color)

    arrow_image_path = os.path.join("img", trend_arrow + ".png")
    arrow_image = Image.open(arrow_image_path)

    if arrow_image.mode != "RGBA":
        arrow_image = arrow_image.convert("RGBA")

    color_image = Image.new("RGBA", arrow_image.size, text_color)
    mask = arrow_image.split()[3].point(lambda x: x > 0 and 255)
    color_image.putalpha(mask)
    arrow_image = Image.alpha_composite(arrow_image, color_image)

    available_height = icon_size[1] - text_height
    arrow_height = min(available_height, arrow_image.size[1])
    arrow_width = int(arrow_height * arrow_image.size[0] / arrow_image.size[1])

    arrow_image = arrow_image.resize((arrow_width, arrow_height), 0)
    arrow_x = (icon_size[0] - arrow_width) // 2
    arrow_y = icon_size[1] - arrow_height

    image.paste(arrow_image, (arrow_x, arrow_y), arrow_image)

    return image


def get_current_blood_glucose():
    try:
        if nightscout_url:
            api_url = nightscout_url + "api/v1/entries.json?count=1"
            response = requests.get(api_url)
            response.raise_for_status()
            entries = response.json()
            if entries:
                entry = entries[0]
                blood_glucose = entry.get("sgv")
                trend_arrow = entry.get("direction")
                timestamp = entry.get("date") / 1000
                print(timestamp, blood_glucose)
                return blood_glucose, trend_arrow, timestamp
    except Exception as e:
        print(f"Failed to fetch current blood glucose: {str(e)}")
    return None, None, None


def update_icon(icon):
    global last_blood_glucose, last_trend_arrow, last_value_timestamp, delta

    while True:
        blood_glucose, trend_arrow, timestamp = get_current_blood_glucose()

        if blood_glucose is not None and trend_arrow is not None and timestamp is not None:
            if blood_glucose != last_blood_glucose or trend_arrow != last_trend_arrow or timestamp != last_value_timestamp:
                delta = blood_glucose - last_blood_glucose
                last_blood_glucose = blood_glucose
                last_trend_arrow = trend_arrow
                last_value_timestamp = timestamp
                icon.icon = create_icon_image(last_blood_glucose, last_trend_arrow)
                icon.update_menu()

        if last_value_timestamp is not None:
            next_update_time = last_value_timestamp + 300
            current_time = time.time()
            time_until_update = next_update_time - current_time

            if time_until_update < -100:
                sleep_duration = 300
            else:
                sleep_duration = max(10, time_until_update)

            time.sleep(sleep_duration)


def update_icon_once():
    global nightscout_url, last_blood_glucose, last_trend_arrow, last_value_timestamp, icon

    blood_glucose, trend_arrow, timestamp = get_current_blood_glucose()
    last_blood_glucose = blood_glucose
    last_trend_arrow = trend_arrow
    last_value_timestamp = timestamp
    icon.icon = create_icon_image(last_blood_glucose, last_trend_arrow)
    icon.update_menu()
    

def get_tooltip():
    global last_value_timestamp, delta
    if last_value_timestamp:
        diff_minutes = int((time.time() - last_value_timestamp) / 60)
        if diff_minutes == 1:
            m = "minute"
        else:
            m = "minutes"
        text_time = f"{diff_minutes} {m} ago"
    else:
        text_time = ""
    if delta:
        if delta >= 0:
            sign = "+"
        else:
            sign = "-"
        text_delta = f"\n{sign}{abs(delta)} {unit}"
    else:
        text_delta = ""
    return text_time + text_delta


def update_title(icon):
    global last_value_timestamp

    update_title_timestamp = last_value_timestamp
    last_timestamp = last_value_timestamp

    while True:
        if last_value_timestamp != last_timestamp:
            update_title_timestamp = last_value_timestamp
            last_timestamp = last_value_timestamp
            icon.title = get_tooltip()
            print('title updated')

        current_time = time.time()
        elapsed_time = current_time - update_title_timestamp

        if elapsed_time >= 60:
            icon.title = get_tooltip()
            update_title_timestamp += 60
            print('title updated')

        time.sleep(1)


def on_exit(icon, item):
    icon.stop()
    os._exit(0)


def is_dark_color(color):
    # Convert the color to RGB values
    r, g, b = ImageColor.getrgb(color)

    # Calculate the relative luminance using the formula
    # L = 0.2126 * R + 0.7152 * G + 0.0722 * B
    luminance = 0.2126 * r + 0.7152 * g + 0.0722 * b

    # Define a threshold value to determine darkness or lightness
    threshold = 128

    # Return True if the luminance is below the threshold (dark color), False otherwise (light color)
    return luminance < threshold



def adjust_range():
    # Create a tkinter root window for the dialog
    root = tk.Tk()
    root.withdraw()  # Hide the root window

    # Create a custom dialog window
    dialog = tk.Toplevel(root)
    dialog.title("Adjust Range")
    dialog.geometry("350x150")

    # Retrieve the current target range and colors
    current_lower_range, current_upper_range = target_range
    current_lower_color, current_upper_color = target_colors

    # Create the title label
    tk.Label(dialog, text="Glucose Target Range", font=("Arial", 12, "bold")).grid(row=0, column=0, columnspan=4)

    def get_font_color(color):
        # Check if the color is dark or light and return the appropriate font color
        if is_dark_color(color):
            return "white"
        else:
            return "black"

    # Create input fields for target range thresholds
    tk.Label(dialog, text="Low:", width=8, anchor="e").grid(row=1, column=0, sticky=tk.E)
    lower_range_entry = tk.Entry(dialog, width=10)
    lower_range_entry.insert(tk.END, current_lower_range)
    lower_range_entry.grid(row=1, column=1, padx=(0, 5), sticky=tk.W)
    lower_range_entry.config(bg=target_colors[0])
    lower_range_entry.config(fg=get_font_color(target_colors[0]))

    tk.Label(dialog, text="High:", width=8, anchor="e").grid(row=1, column=2, sticky=tk.E)
    upper_range_entry = tk.Entry(dialog, width=10)
    upper_range_entry.insert(tk.END, current_upper_range)
    upper_range_entry.grid(row=1, column=3, padx=(0, 5), sticky=tk.W)
    upper_range_entry.config(bg=target_colors[1])
    upper_range_entry.config(fg=get_font_color(target_colors[1]))


    # Create a color picker button for lower range color
    def pick_lower_color():
        _, color = colorchooser.askcolor(title="Select Color for Lower Range Threshold")
        if color:
            # lower_color_picker_button.config(bg=color)
            lower_range_entry.config(bg=color)
            lower_range_entry.config(fg=get_font_color(color))
        return color

    lower_color_picker_button = tk.Button(dialog, text="Pick Color", command=pick_lower_color)
    lower_color_picker_button.grid(row=2, column=1, pady=(0, 10), padx=(0, 5), sticky=tk.W)


    # Create a color picker button for upper range color
    def pick_upper_color():
        _, color = colorchooser.askcolor(title="Select Color for Upper Range Threshold")
        if color:
            # upper_color_picker_button.config(bg=color)
            upper_range_entry.config(bg=color)
            upper_range_entry.config(fg=get_font_color(color))

    upper_color_picker_button = tk.Button(dialog, text="Pick Color", command=pick_upper_color)
    upper_color_picker_button.grid(row=2, column=3, pady=(0, 10), padx=(0, 5), sticky=tk.W)

    # Add a button to save the settings
    def save_settings():
        lower_range = lower_range_entry.get()
        upper_range = upper_range_entry.get()
        lower_color = lower_range_entry.cget("bg")
        upper_color = upper_range_entry.cget("bg")

        # Update the target range thresholds and colors
        update_target_range((lower_range, upper_range), (lower_color, upper_color))

        # Call the update_icon_once() function to update the icon immediately
        update_icon_once()

        dialog.destroy()

    save_button = tk.Button(dialog, text="Save", command=save_settings)
    save_button.grid(row=3, column=0, columnspan=4, pady=(10, 0))

    # Center the dialog window on the screen
    window_width = dialog.winfo_reqwidth()
    window_height = dialog.winfo_reqheight()
    position_right = int(dialog.winfo_screenwidth() / 2 - window_width / 2)
    position_down = int(dialog.winfo_screenheight() / 2 - window_height / 2)
    dialog.geometry(f"+{position_right}+{position_down}")

    # Run the dialog window
    dialog.mainloop()



def update_target_range(selected_range, selected_colors):
    global target_range, target_colors
    target_range = selected_range
    target_colors = selected_colors



def open_adjust_url():
    global nightscout_url, last_blood_glucose, last_trend_arrow, last_value_timestamp

    # Create a new thread to run the settings dialog
    settings_thread = threading.Thread(target=run_adjust_url_dialog)
    settings_thread.start()


def run_adjust_url_dialog():
    global nightscout_url
    # Create a tkinter root window for the dialog
    root = tk.Tk()
    root.withdraw()  # Hide the root window

    # Prompt the user to enter the Nightscout URL
    new_url = simpledialog.askstring("Nightscout URL", "Enter the Nightscout URL:", parent=root)

    if new_url:
        nightscout_url = new_url
        urls = read_urls_from_file()
        if new_url not in urls:
            save_url_to_file(new_url)



def create_system_tray_icon():
    global last_blood_glucose, last_trend_arrow, last_value_timestamp


    # Create a menu with an initial placeholder icon
    menu = (
        pystray.MenuItem("Nightscout URL", open_adjust_url),
        pystray.MenuItem("Target Range", adjust_range),
        pystray.MenuItem("Exit", on_exit),
    )

    # Create the icon with the menu
    default_icon_image = Image.new("RGBA", (64, 64), (255, 255, 255, 0))
    icon = pystray.Icon("cgm_icon", default_icon_image, "cgm_icon", menu)

    # Set the exit handler
    icon.hook = on_exit

    return icon


def initialize_url():
    global nightscout_url
    urls = read_urls_from_file()
    if not urls:
        open_adjust_url()
    else:
        nightscout_url = urls[0]



if __name__ == "__main__":
    icon = create_system_tray_icon()
    initialize_url()

    while not nightscout_url:
        time.sleep(1)

    update_icon_once()

    # Run the icon update loop in a separate thread
    update_icon_thread = threading.Thread(target=update_icon, args=(icon,))
    update_icon_thread.start()

    # Run the title update loop in a separate thread
    update_title_thread = threading.Thread(target=update_title, args=(icon,))
    update_title_thread.start()

    # Run the system tray application
    icon.run()