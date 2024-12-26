from flask import Flask, render_template, request
import threading
import cv2
import mediapipe as mp
import pyautogui
import math
import pycaw
from ctypes import cast, POINTER
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

from comtypes import CLSCTX_ALL

import pythoncom
import pyttsx3
import speech_recognition as sr
import os
import webbrowser
import datetime


app = Flask(__name__)

# Global flag to stop the hand control thread
exit_flag = False

# Initialize audio control
devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))

# Define the hand control function with threading support
def hand_control_thread():
    global exit_flag
    cap = cv2.VideoCapture(0)
    if not cap.isOpened():
        print("Error: Could not access the webcam.")
        return
    
    hand_detector = mp.solutions.hands.Hands()
    drawing_utils = mp.solutions.drawing_utils
    screen_width, screen_height = pyautogui.size()  # Get the screen size
    plocx, plocy = 0, 0
    clocx, clocy = 0, 0
    smoothening = 3

    index_x, index_y = 0, 0  # Initialize index_x and index_y variables
    thumb_x, thumb_y = 0, 0  # Initialize thumb_x and thumb_y variables

    while True:
        if exit_flag:
            print("Exit flag set, closing hand control...")
            break  # Exit the loop to stop the hand control thread
        
        ret, frame = cap.read()
        if not ret:
            print("Failed to capture frame. Exiting...")
            break
        
        frame = cv2.flip(frame, 1)
        frame_height, frame_width, _ = frame.shape
        rgb_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        output = hand_detector.process(rgb_frame)
        hands = output.multi_hand_landmarks
        
        if hands:
            for hand in hands:
                drawing_utils.draw_landmarks(frame, hand, mp.solutions.hands.HAND_CONNECTIONS)
                landmarks = hand.landmark
                
                # Access thumb and index finger coordinates and scale to screen size
                thumb_x = int(landmarks[4].x * frame_width)  # Thumb tip
                thumb_y = int(landmarks[4].y * frame_height)
                index_x = int(landmarks[8].x * frame_width)  # Index finger tip
                index_y = int(landmarks[8].y * frame_height)

                # Calculate distance between thumb and index finger
                distance = math.sqrt((thumb_x - index_x) ** 2 + (thumb_y - index_y) ** 2)
                
                # Define a threshold for the pinch gesture
                if distance < 30:  # Adjust this value as needed for sensitivity
                    cv2.putText(frame, "Click!", (50, 50), cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 255, 0), 2)
                    pyautogui.click()  # Perform a click
                    pyautogui.sleep(1)  # Prevent multiple clicks; adjust the delay as needed

                # Smooth mouse movement by scaling to screen size
                clocx = (index_x / frame_width) * screen_width  # Map index finger x to screen width
                clocy = (index_y / frame_height) * screen_height  # Map index finger y to screen height
                clocx = plocx + (clocx - plocx) / smoothening
                clocy = plocy + (clocy - plocy) / smoothening
                pyautogui.moveTo(clocx, clocy)
                plocx, plocy = clocx, clocy

        # Show the frame in fullscreen mode (if needed)
        cv2.imshow('Virtual Mouse', frame)
        
        key = cv2.waitKey(1)
        if key == ord('q'):
            break

    cap.release()
    cv2.destroyAllWindows()




# Function to start hand control in a separate thread
def hand_control():
    global exit_flag
    exit_flag = False  # Reset the exit flag when hand control starts
    threading.Thread(target=hand_control_thread).start()
    return "Hand Control activated!"

# Function to stop the hand control thread by setting the exit flag
def exit_program():
    global exit_flag
    exit_flag = True  # Set the exit flag to stop the hand control loop
    return "Exiting the program..."



#eye control
# Global flag to stop the eye control thread
exit_flag = False
def eye_control_thread():
    global exit_flag
    screen_w, screen_h = pyautogui.size()  # Get the screen dimensions
    cam = cv2.VideoCapture(0)
    face_mesh = mp.solutions.face_mesh.FaceMesh(refine_landmarks=True)

    def find_landmarks_and_click(landmarks, frame_w, frame_h):
        for id, landmark in enumerate(landmarks[474:478]):  # Detect eye landmarks
            x = int(landmark.x * frame_w)
            y = int(landmark.y * frame_h)
            cv2.circle(frame, (x, y), 3, (0, 255, 0))  # Draw eye landmarks
            
            if id == 1:  # The middle landmark between the eyes
                screen_x = (landmark.x * screen_w)  # Scale to screen width
                screen_y = (landmark.y * screen_h)  # Scale to screen height
                pyautogui.moveTo(screen_x, screen_y)  # Move the cursor to the new position

        # Detect if the eyes are closed to click
        left_top = landmarks[145]  # Upper point of the left eye
        left_bottom = landmarks[159]  # Lower point of the left eye

        # Convert normalized coordinates to pixel values
        left_top_y = int(left_top.y * frame_h)
        left_bottom_y = int(left_bottom.y * frame_h)

        # Calculate vertical distance between the two points
        left_eye_dist = abs(left_top_y - left_bottom_y)

        # Set a threshold for detecting eye closure (tune this as per your setup)
        threshold = 5  # Adjust based on the size of your camera feed

        # If the distance is below the threshold, it's a blink (click)
        if left_eye_dist < threshold:
            cv2.putText(frame, "Click Detected!", (50, 50), cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 255, 0), 2)
            pyautogui.click()  # Perform the click action
            pyautogui.sleep(1)  # Delay to avoid multiple clicks

    while True:
        if exit_flag:  # Check if the exit flag is set
            print("Exit flag set, closing eye control...")
            break  # Exit the loop to stop the eye control thread
        
        ret, frame = cam.read()
        if not ret:
            break
        frame = cv2.flip(frame, 1)  # Flip the frame horizontally
        rgb_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        output = face_mesh.process(rgb_frame)  # Process the frame for facial landmarks
        landmark_points = output.multi_face_landmarks
        
        if landmark_points:
            landmarks = landmark_points[0].landmark  # Get the landmarks
            frame_h, frame_w, _ = frame.shape  # Get the frame dimensions
            find_landmarks_and_click(landmarks, frame_w, frame_h)  # Call the function to find landmarks and click
        
        cv2.imshow('Eye Controlled Mouse', frame)  # Show the video feed
        
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break  # Exit on 'q' key press
            
    cam.release()  # Release the camera
    cv2.destroyAllWindows()  # Close all OpenCV windows

# Function to start eye control in a separate thread
def eye_control():
    global exit_flag
    exit_flag = False  # Reset the exit flag when eye control starts
    threading.Thread(target=eye_control_thread).start()
    return "Eye Control activated!"



voice_control_flag = False

def assist():
    recognizer = sr.Recognizer()
    engine = pyttsx3.init()

    recognizer.energy_threshold = 4000
    voices = engine.getProperty('voices')
    engine.setProperty('voice', voices[1].id)  # Use female voice

    def speak(text):
        engine.say(text)
        engine.runAndWait()

    def listen():
        with sr.Microphone() as source:
            print("Listening...")
            audio = recognizer.listen(source)

        try:
            print("Recognizing...")
            command = recognizer.recognize_google(audio)
            print("You said:", command)
            return command.lower()
        except sr.UnknownValueError:
            print("Sorry, I couldn't understand what you said.")
            return ""
        except sr.RequestError:
            print("Sorry, I encountered an error while processing your request.")
            return ""

    def execute_command(command):
        if "hello" in command:
            speak("Hello! How can I assist you today?")
        elif "time" in command:
            current_time = datetime.datetime.now().strftime("%I:%M %p")
            speak(f"The current time is {current_time}")
        elif "date" in command:
            current_date = datetime.datetime.now().strftime("%B %d, %Y")
            speak(f"Today's date is {current_date}")
        elif "search for" in command:
            # Extract the search query and open it in the default web browser
            search_query = command.replace("search for", "").strip()
            url = f"https://www.google.com/search?q={search_query}"
            webbrowser.open(url)
            speak(f"Searching for {search_query} on Google.")
        elif "exit" in command:
            speak("Goodbye!")
            return False  # Stop voice control
        else:
            speak("I'm sorry, I don't understand that command.")
        return True

    speak("Voice control activated! How can I help you?")
    while True:
        command = listen()
        if command and not execute_command(command):
            break

def voice_control_thread():
    # Initialize COM for this thread
    pythoncom.CoInitialize()
    
    # Call the assist function
    assist()

# Function to start voice control in a separate thread
def start_voice_control():
    threading.Thread(target=voice_control_thread).start()
    return "Voice Control activated!"


# Flask route for the main page
@app.route('/')
def index():
    return render_template('index.html')

# Flask route to handle button clicks
@app.route('/', methods=['POST'])
def handle_button_click():
    action = request.form['action']
    if action == "Hand Control":
        result = hand_control()
    elif action == "Exit Program":
        result = exit_program()# Exit the hand control program
    elif action == "Eye Control":
        result = eye_control() 
    elif action == 'Voice Control':
            result = start_voice_control() 
    else:
        result = "Unknown action"

    return f"<h1>{result}</h1><a href='/'>Go back</a>"

if __name__ == '__main__':
    app.run(debug=True)
