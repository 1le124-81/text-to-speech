# text-to-speech
a simple local tts for discord and other call stuffs

When run, it opens a tkinter window with a textbox, 2 sliders, and a button. The sliders can be adjusted to control the speed of playback and the volume, while the button speaks the text that is input into the textbox. 

Uses windows SAPI for tts and pythomcom and win32com.client to relay text to it. Uses threading to ensure ui does not freeze when the tts is speaking.

The generated voice is then relayed into a voice call through VB-Cable (requires separate downloading). youll also have to change ur microphone (input device) to VB-Cable.

autodeletes text in textbox after speaking, can also use enter instead of pressing the speak button. Only compatible with windows pcs/laptops i think.

<img width="499" height="430" alt="image" src="https://github.com/user-attachments/assets/45eb772e-5ba4-44a3-8cf4-0ea5c66e154c" />
