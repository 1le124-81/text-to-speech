#import libraries
import tkinter as tk #for gui
from tkinter import messagebox
import win32com.client #lets us communicate with com, including sapi.
import threading #prevents freezing the gui by running multiple threads
import pythoncom #com initialization for threads

#finds cable
def find_cable():
    voice = win32com.client.Dispatch("SAPI.SpVoice") #defines com of sapi.spvoice

    outputs = voice.GetAudioOutputs("", "") #gets available audio outputs, strings empty bcz we dont care about the category or subcategory of the output

    #loops through all available audio outputs and checks if the name contains "CABLE Input"
    for i in range(outputs.Count):
        device = outputs.Item(i)
        name = device.GetDescription()

        print("Found audio device:", name)

        if "CABLE Input" in name:
            return device #returns the virtual cable

    return None #return none if no cable input was found

#speaking
def speak_text(text, speed, volume):
    pythoncom.CoInitialize() #initializes com for the thread

    try:
        voice = win32com.client.Dispatch("SAPI.SpVoice") #defines com of sapi.spvoice again

        cable = find_cable() #defines cable as output of function find_cable

        if cable is None: #tk shows popup when no cable input is found, and returns.
            messagebox.showerror(
                "VB-CABLE not found",
                "Couldn't find CABLE Input."
            )
            return

        #sets the audio output of sapi to cable so the voice is spoken through it
        voice.AudioOutput = cable

        #changes voice speed and volume based on user input
        voice.Rate = speed
        voice.Volume = volume

        #uses sapi to actually speak the text
        voice.Speak(text)

    finally:
        pythoncom.CoUninitialize() #closes up com after runnign


#gui input to speak
def speak():
    text = text_box.get("1.0", tk.END).strip() #gets text from the text box, starting at line 1, character 0 until end of the text box, removes whitespace

    #if no text do nothing
    if not text:
        return

    #clears textbox
    text_box.delete("1.0", tk.END)

    #sets speed and volume to value set by sliders
    speed = int(speed_slider.get())
    volume = int(volume_slider.get())

    #creates new thread that runs speak_text function with the text as an argument (comma signifies one argument), sets as daemon so it closes when the main program closes 
    threading.Thread(
        target=speak_text,
        args=(text, speed, volume),
        daemon=True
    ).start()


#runs speak when enter is pressed, returns "break" to prevent new line from being added to the text box
def enter_pressed(event):
    speak()
    return "break"

#sets up gui
root = tk.Tk()

#title and size of the tts gui, makes it not resizable
root.title("Discord TTS")
root.geometry("500x400")
root.resizable(False, False)

#creates a label at top of window
label = tk.Label(
    root,
    text="enter text here"
)

label.pack(pady=(15, 5)) #sets label psoition

#adds textbox to the tk window, sets width and height, and font
text_box = tk.Text(
    root,
    width=55,
    height=8,
    font=("Segoe UI", 11)
)

text_box.pack() #adds to window using pack layout

#volume label and slider
speed_label = tk.Label(root, text="Speed") #put volumelabel text in root (normal window) with speed text
speed_label.pack() #adds to window using pack layout

speed_slider = tk.Scale(  #scale = slider module for tk
    root, #put in normal window
    from_=-10, #slider range from -10 to 10
    to=10,
    orient=tk.HORIZONTAL, #horizontal orientation
    length=300 #length of slider
)

speed_slider.set(0) #default
speed_slider.pack() #adds to window using pack layout

#volume label and slider
volume_label = tk.Label(root, text="Volume")  #put volumelabel text in root (normal window) with speed text
volume_label.pack() #adds to window using pack layout

volume_slider = tk.Scale( #scale = slider module for tk
    root, #put in normal window
    from_=0, #slider range from 0 to 100
    to=100,
    orient=tk.HORIZONTAL, #horizontal orientation
    length=300 #length of slider
)

volume_slider.set(100) #default
volume_slider.pack() #adds to window using pack layout

#creates speak button, sets text, command to run when clicked, width and height
button = tk.Button(
    root,
    text="SPEAK",
    command=speak,
    width=15,
    height=1
)

#button position
button.pack(pady=10)

#when return is pressed in the text box, run enter_pressed function
text_box.bind("<Return>", enter_pressed)

#focus on the text box when program starts
text_box.focus()

#starts gui
root.mainloop()