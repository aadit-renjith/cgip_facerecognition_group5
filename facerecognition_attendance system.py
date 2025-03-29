import cv2
import numpy as np
import face_recognition
import datetime
import pandas as pd
import os
import pickle

# Initialize webcam
cam = cv2.VideoCapture(0)

# Load known faces if available
pickle_file = 'Entered_faces_data.pkl'
known_faces_encodings = []
known_faces_names = []

if os.path.exists(pickle_file) and os.path.getsize(pickle_file) > 0:
    with open(pickle_file, 'rb') as file:
        known_faces_encodings, known_faces_names = pickle.load(file)

attendance = {}

# Add new student
new_student = input("New Student(yes/no): ").strip().lower()
if new_student == "yes":
    student_name = input("Name: ").strip()
    print("Please look at the camera...")

    for i in range(20):
        ret, frame = cam.read()
        if not ret:
            print("Failed to capture image. Exiting...")
            break
        cv2.imshow('Capturing Image', frame)
        if cv2.waitKey(1) == ord('q'):
            break

    image_path = f"{student_name}.jpg"
    cv2.imwrite(image_path, frame)

    image = face_recognition.load_image_file(image_path)
    encodings = face_recognition.face_encodings(image)
    
    if encodings:
        encoding = encodings[0]
        known_faces_encodings.append(encoding)
        known_faces_names.append(student_name)
        with open(pickle_file, 'wb') as file:
            pickle.dump((known_faces_encodings, known_faces_names), file)
        print(f"New student {student_name} added.")
    else:
        print("No face detected, try again.")

while True:
    ret, frame = cam.read()
    if not ret:
        break

    rgb_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
    face_locations = face_recognition.face_locations(rgb_frame)
    face_encodings = face_recognition.face_encodings(rgb_frame, face_locations)

    for face_encoding, face_location in zip(face_encodings, face_locations):
        matches = face_recognition.compare_faces(known_faces_encodings, face_encoding)
        name = "Unknown"
        
        if True in matches:
            match_indexes = [i for i, match in enumerate(matches) if match]
            name = known_faces_names[match_indexes[0]]
            if name not in attendance:
                attendance[name] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        top, right, bottom, left = face_location
        cv2.rectangle(frame, (left, top), (right, bottom), (0, 255, 0), 2)
        cv2.putText(frame, name, (left, top - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.75, (0, 255, 0), 2)
    
    cv2.imshow('Camera', frame)
    if cv2.waitKey(1) == ord('q'):
        break

# Save attendance to an Excel file
if attendance:
    df = pd.DataFrame([{"Name": name, "Timestamp": time} for name, time in attendance.items()])
    df.to_excel(r"C:\Users\aadit\Documents\Codes\CGIP_MICROPROJECT_S6\attendance_log.xlsx", index=False)
    print("Attendance saved to attendance_log.xlsx")

cam.release()
cv2.destroyAllWindows()