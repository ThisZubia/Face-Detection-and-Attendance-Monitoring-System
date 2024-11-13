import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from datetime import datetime
import cv2
from PIL import Image, ImageTk
import os
import numpy as np
from Emplyee_code import EmployeeManagementSystem

class FaceDetectionAttendanceSystem:
    def __init__(self, root):
        self.root = root
        self.root.title("Face Detection Attendance System")
        self.root.geometry("1030x700")
        self.root.configure(bg="white")

        # Initialize camera and face detection
        self.cap = cv2.VideoCapture(0)
        self.face_cascade = cv2.CascadeClassifier(cv2.data.haarcascades + 'haarcascade_frontalface_default.xml')

        # Initialize the LBPH face recognizer
        self.face_recognizer = cv2.face.LBPHFaceRecognizer_create()

        # Employee database
        self.employee_data = {"Emp001": "John Doe"}  # Initially, only one employee is added

        # Setup UI components
        self.create_ui()

        # Ensure 'Data' folder exists for storing face images and trained model
        if not os.path.exists("Data"):
            os.makedirs("Data")

        # Load the trained model if it exists
        if os.path.exists("Data/face_recognizer.yml"):
            self.face_recognizer.read("Data/face_recognizer.yml")

        # Flag to track if an employee has logged in
        self.logged_in_employees = {}

        # Attendance log storage
        self.attendance_log = []

    def create_ui(self):
        # Attendance Panel
        attendance_frame = tk.LabelFrame(self.root, text="Attendance Log", padx=10, pady=10, bg="white", fg="black", font=("Arial", 12))
        attendance_frame.place(x=20, y=20, width=450, height=515)

        self.attendance_tree = ttk.Treeview(attendance_frame, columns=("Employee ID", "Name", "Date", "Time", "Status"), show="headings")
        self.attendance_tree.heading("Employee ID", text="Employee ID")
        self.attendance_tree.heading("Name", text="Name")
        self.attendance_tree.heading("Date", text="Date")
        self.attendance_tree.heading("Time", text="Time")
        self.attendance_tree.heading("Status", text="Status")
        self.attendance_tree.column("Employee ID", width=80, anchor="center")
        self.attendance_tree.column("Name", width=120, anchor="center")
        self.attendance_tree.column("Date", width=80, anchor="center")
        self.attendance_tree.column("Time", width=80, anchor="center")
        self.attendance_tree.column("Status", width=80, anchor="center")
        self.attendance_tree.pack(fill="both", expand=True)

        # Face Recognition Panel
        face_recognition_frame = tk.LabelFrame(self.root, text="Face Recognition", padx=10, pady=10, bg="white", fg="black", font=("Arial", 12))
        face_recognition_frame.place(x=500, y=20, width=515, height=515)

        # Display area for live video feed
        self.video_label = tk.Label(face_recognition_frame, text="Live Feed", bg="black", width=450, height=400)
        self.video_label.pack(pady=10)

        # Control Buttons for Attendance
        button_frame = tk.Frame(face_recognition_frame, bg="white")
        button_frame.pack(pady=10)

        self.capture_in_button = tk.Button(button_frame, text="Capture & Log IN", command=lambda: self.capture("IN"), width=15, font=("Arial", 10))
        self.capture_in_button.grid(row=0, column=0, padx=10)

        self.capture_out_button = tk.Button(button_frame, text="Capture & Log OUT", command=lambda: self.capture("OUT"), width=15, font=("Arial", 10))
        self.capture_out_button.grid(row=0, column=1, padx=10)

        # Status Bar
        self.status_bar = tk.Label(self.root, text="System Messages", bd=1, relief=tk.SUNKEN, anchor=tk.W, font=("Arial", 10), bg="white")
        self.status_bar.place(x=20, y=550, width=997, height=30)

        # Additional Control Buttons below Status Bar
        controls_frame = tk.Frame(self.root, bg="white")
        controls_frame.place(x=20, y=600, width=997, height=50)

        edit_button = tk.Button(controls_frame, text="Edit", command=self.open_employee_management, width=15, font=("Arial", 10))  # Edit button's functionality removed
        edit_button.grid(row=0, column=1, padx=10)

        view_report_button = tk.Button(controls_frame, text="View Report", command=self.view_report, width=15, font=("Arial", 10))
        view_report_button.grid(row=0, column=2, padx=10)

        # Start video stream
        self.update_video_stream()

    def add_employee(self):
        employee_id = simpledialog.askstring("Input", "Enter Employee ID:")
        name = simpledialog.askstring("Input", "Enter Employee Name:")

        if employee_id and name:
            # Add the new employee to the dictionary
            self.employee_data[employee_id] = name
            captured_faces = []
            count = 0

            while count < 10:
                ret, frame = self.cap.read()
                if not ret:
                    continue
                
                gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
                faces = self.face_cascade.detectMultiScale(gray, scaleFactor=1.1, minNeighbors=5)
                
                for (x, y, w, h) in faces:
                    count += 1
                    face_img = gray[y:y+h, x:x+w]
                    captured_faces.append(face_img)
                    # Save each captured face
                    cv2.imwrite(f"Data/{employee_id}_{count}.jpg", face_img)
                    
                    # Display feedback for capture
                    cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 255, 0), 2)
                    cv2.putText(frame, f"Capturing image {count}", (x, y - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 255, 0), 2)
                    cv2.imshow("Capturing Face", frame)
                    cv2.waitKey(100)
                    
                if count >= 10:
                    break
            
            cv2.destroyWindow("Capturing Face")
            self.train_face_recognizer()  # Retrain recognizer with the new data
            messagebox.showinfo("Success", f"Employee {name} added and trained successfully!")

        else:
            messagebox.showerror("Error", "Employee ID and Name are required!")

    def delete_employee(self):
        employee_id = simpledialog.askstring("Input", "Enter Employee ID to Delete:")
        if employee_id in self.employee_data:
            del self.employee_data[employee_id]
            # Remove saved face images
            for file in os.listdir("Data"):
                if file.startswith(f"{employee_id}_"):
                    os.remove(f"Data/{file}")
            messagebox.showinfo("Success", f"Employee ID {employee_id} deleted successfully!")
            self.train_face_recognizer()  # Retrain recognizer after deletion
        else:
            messagebox.showerror("Error", "Employee not found!")

    def train_face_recognizer(self):
        # Train the face recognizer using all stored images
        faces = []
        labels = []

        for file in os.listdir("Data"):
            if file.endswith(".jpg"):
                employee_id = file.split("_")[0]  # Extract the employee ID part (Emp001, Emp002, etc.)
                if employee_id.startswith("Emp"):
                    label = int(employee_id[3:])  # Extract label number from 'EmpXXX'
                    img_path = os.path.join("Data", file)
                    img = cv2.imread(img_path, cv2.IMREAD_GRAYSCALE)
                    faces.append(img)
                    labels.append(label)

        if faces and labels:
            labels = np.array(labels, dtype=np.int32)
            self.face_recognizer.train(faces, labels)
            self.face_recognizer.save("Data/face_recognizer.yml")  # Save trained model
        else:
            messagebox.showerror("Error", "No face data to train the recognizer!")

    def capture(self, status):
        # Capture image and log attendance without checking if model is trained
        ret, frame = self.cap.read()
        if not ret:
            messagebox.showerror("Error", "Failed to capture image!")
            return
        
        # Save the captured image (without face detection)
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        image_filename = f"Data/captured_{timestamp}.jpg"
        cv2.imwrite(image_filename, frame)

        # Log the attendance
        current_time = datetime.now().strftime("%H:%M:%S")
        current_date = datetime.now().strftime("%Y-%m-%d")

        # Record login or logout status
        status_str = "Logged IN" if status == "IN" else "Logged OUT"
        self.attendance_log.append({
            "Employee ID": "Unknown", "Name": "Unknown", "Date": current_date, "Time": current_time, "Status": status_str
        })
        self.attendance_tree.insert("", "end", values=("Unknown", "Unknown", current_date, current_time, status_str))
        self.status_bar.config(text=f"Image captured, {status_str} at {current_time}")

    def view_report(self):
        # Show full attendance report
        report_window = tk.Toplevel(self.root)
        report_window.title("Monthly Attendance Report")
        report_window.geometry("800x600")

        report_tree = ttk.Treeview(report_window, columns=("Employee ID", "Name", "Date", "Time", "Status"), show="headings")
        report_tree.heading("Employee ID", text="Employee ID")
        report_tree.heading("Name", text="Name")
        report_tree.heading("Date", text="Date")
        report_tree.heading("Time", text="Time")
        report_tree.heading("Status", text="Status")
        report_tree.column("Employee ID", width=80, anchor="center")
        report_tree.column("Name", width=120, anchor="center")
        report_tree.column("Date", width=80, anchor="center")
        report_tree.column("Time", width=80, anchor="center")
        report_tree.column("Status", width=80, anchor="center")
        report_tree.pack(fill="both", expand=True)

        # Populate the report with attendance log data
        for record in self.attendance_log:
            report_tree.insert("", "end", values=(record["Employee ID"], record["Name"], record["Date"], record["Time"], record["Status"]))
        
        # Close button
        close_button = tk.Button(report_window, text="Close", command=report_window.destroy)
        close_button.pack(pady=10)

    def update_video_stream(self):
        # Capture frame from camera and update the GUI
        ret, frame = self.cap.read()
        if ret:
            cv2image = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            img = Image.fromarray(cv2image)
            img_tk = ImageTk.PhotoImage(image=img)
            self.video_label.img_tk = img_tk
            self.video_label.config(image=img_tk)

        self.root.after(10, self.update_video_stream)

    def open_employee_management(self):
        # Open the Employee Management System
        emp_manager = EmployeeManagementSystem(self)
        emp_manager.run()

    def open_employee_management(self):
        new_window = tk.Toplevel(self.root)
        app = EmployeeManagementSystem(new_window)  # Create an instance of EmployeeManagementSystem
        new_window.mainloop()
if __name__ == "__main__":
    root = tk.Tk()
    app = FaceDetectionAttendanceSystem(root)
    root.mainloop()
