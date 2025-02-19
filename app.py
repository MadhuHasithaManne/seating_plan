from flask import Flask, render_template, request, session, redirect, url_for, send_from_directory
import pandas as pd
import zipfile
import os
import uuid

app = Flask(__name__)
app.secret_key = "supersecretkey12345"  # Mandatory for session handling

# Directory for storing generated files
OUTPUT_DIR = os.path.join(os.getcwd(), 'static', 'output_files')
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Pass Python's zip function to Jinja2 templates
app.jinja_env.globals.update(zip=zip)

# Directory for storing temporary files
TEMP_DIR = os.path.join(os.getcwd(), 'temp')
os.makedirs(TEMP_DIR, exist_ok=True)

@app.before_request
def make_session_non_permanent():
    session.permanent = False

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        try:
            # Input from the user
            num_departments = int(request.form["num_departments"])
            department_names = request.form.getlist("department_names[]")
            subject_codes = request.form.getlist("subject_codes[]")
            uploaded_file = request.files["roll_numbers_file"]

            if uploaded_file.filename == "":
                return "Error: No file uploaded."

            # Load roll numbers and department data
            try:
                df = pd.read_excel(uploaded_file)
                print("File uploaded and read successfully.")
            except Exception as e:
                return f"Error: Unable to read the uploaded file. {str(e)}"

            # Validate input
            if len(department_names) != num_departments or len(subject_codes) != num_departments:
                return "Error: Number of departments or subject codes doesn't match the provided details."

            if not set(department_names).issubset(set(df["Department"].unique())):
                return "Error: Department names do not match with the uploaded file."

            # Map subject codes to departments
            department_subject_map = dict(zip(department_names, subject_codes))

            # Filter data to include only relevant departments
            df = df[df["Department"].isin(department_names)]

            # Save processed data to a temporary file
            session_id = str(uuid.uuid4())
            temp_file_path = os.path.join(TEMP_DIR, f"{session_id}.csv")
            df.to_csv(temp_file_path, index=False)

            # Save session metadata
            session["session_id"] = session_id
            session["num_departments"] = num_departments
            session["department_subject_map"] = department_subject_map

            print("Input validation successful and data stored.")
            return render_template("index.html", buttons_visible=True)

        except Exception as e:
            print(f"Error during form submission: {str(e)}")
            return "An error occurred while processing your request. Please try again."

    return render_template("index.html", buttons_visible=False)
@app.route("/seating_plan")
def seating_plan():
    try:
        # Retrieve session metadata
        session_id = session.get("session_id")
        department_subject_map = session.get("department_subject_map")

        if not (session_id and department_subject_map):
            return "Error: Missing data. Please submit the form first."

        # Load the CSV file containing student data
        temp_file_path = os.path.join(TEMP_DIR, f"{session_id}.csv")
        if not os.path.exists(temp_file_path):
            return "Error: Data file not found. Please submit the form again."

        df = pd.read_csv(temp_file_path)
        df.sort_values(by=["Department", "Roll Number"], inplace=True)

        rooms = []
        room_number = 1
        remaining_students = df.copy()
        remaining_departments = df["Department"].unique().tolist()
        available_departments = remaining_departments.copy()
        remaining_students_dict = {
            dept: {"students": [], "subject_code": department_subject_map.get(dept)}
            for dept in remaining_departments
        }

        # Step 1: Assign students to initial rooms
        while len(remaining_students) > 0:
            room = {"room_number": room_number, "side_a": [], "side_b": []}
            side_a, side_b = [], []

            if available_departments:
                dept_a = available_departments[0]
                dept_a_students = remaining_students[remaining_students["Department"] == dept_a]
                subject_code_a = department_subject_map[dept_a]

                if len(dept_a_students) >= 12:
                    while len(side_a) < 24 and not dept_a_students.empty:
                        student = dept_a_students.iloc[0]
                        side_a.append((student['Department'], student['Roll Number'], subject_code_a))
                        remaining_students = remaining_students[remaining_students["Roll Number"] != student["Roll Number"]]
                        dept_a_students = dept_a_students[dept_a_students["Roll Number"] != student["Roll Number"]]

                    if dept_a_students.empty:
                        available_departments.remove(dept_a)
                else:
                    remaining_students_dict[dept_a]["students"].extend(
                        [(student['Department'], student['Roll Number']) for _, student in dept_a_students.iterrows()]
                    )
                    remaining_students = remaining_students[remaining_students["Department"] != dept_a]
                    available_departments.remove(dept_a)
                    continue  # Skip to the next iteration

            if available_departments:
                for dept_b in available_departments:
                    if department_subject_map[dept_b] != subject_code_a:
                        dept_b_students = remaining_students[remaining_students["Department"] == dept_b]

                        if len(dept_b_students) >= 12:
                            while len(side_b) < 24 and not dept_b_students.empty:
                                student = dept_b_students.iloc[0]
                                side_b.append((student['Department'], student['Roll Number'], department_subject_map[dept_b]))
                                remaining_students = remaining_students[remaining_students["Roll Number"] != student["Roll Number"]]
                                dept_b_students = dept_b_students[dept_b_students["Roll Number"] != student["Roll Number"]]

                            if dept_b_students.empty:
                                available_departments.remove(dept_b)
                            break
                        else:
                            remaining_students_dict[dept_b]["students"].extend(
                                [(student['Department'], student['Roll Number']) for _, student in dept_b_students.iterrows()]
                            )
                            remaining_students = remaining_students[remaining_students["Department"] != dept_b]
                            available_departments.remove(dept_b)

            while len(side_a) < 24:
                side_a.append(("---", "---", None))
            while len(side_b) < 24:
                side_b.append(("---", "---", None))

            room["side_a"] = side_a
            room["side_b"] = side_b
            rooms.append(room)
            room_number += 1

        # Step 2: Assign Unassigned Students to New Rooms
        unassigned_students_list = []
        for dept, data in remaining_students_dict.items():
            for student in data["students"]:
                unassigned_students_list.append((dept, student[1], data["subject_code"]))

        index = 0
        total_unassigned = len(unassigned_students_list)

        while index < total_unassigned:
            room = {"room_number": room_number, "side_a": [], "side_b": []}
            side_a, side_b = [], []

            # Fill Side A first
            while len(side_a) < 24 and index < total_unassigned:
                side_a.append(unassigned_students_list[index])
                index += 1

            # Fill Side B ensuring no same subject codes in the same bench
            for i in range(24):
                if index >= total_unassigned:
                    break
                current_student = unassigned_students_list[index]

                # Ensure subject codes are different for Side A and Side B at the same bench
                if side_a[i][2] is None or side_a[i][2] != current_student[2]:
                    side_b.append(current_student)
                    index += 1
                else:
                    side_b.append(("---", "---", None))  # Keep it empty if subject codes match

            # Fill remaining spots with "---"
            while len(side_a) < 24:
                side_a.append(("---", "---", None))
            while len(side_b) < 24:
                side_b.append(("---", "---", None))

            room["side_a"] = side_a
            room["side_b"] = side_b
            rooms.append(room)
            room_number += 1

        return render_template(
            "seating_plan.html",
            rooms=rooms
        )
    except Exception as e:
        return f"An error occurred: {str(e)}"

@app.route("/generate_attendance_sheets")
def generate_attendance_sheets():
    try:
        session_id = session.get("session_id")
        department_subject_map = session.get("department_subject_map")

        if not (session_id and department_subject_map):
            return "Error: Missing data. Please submit the form first."

        temp_file_path = os.path.join(TEMP_DIR, f"{session_id}.csv")
        if not os.path.exists(temp_file_path):
            return "Error: Data file not found. Please submit the form again."

        df = pd.read_csv(temp_file_path)
        df.sort_values(by=["Department", "Roll Number"], inplace=True)

        rooms = []
        room_number = 1
        remaining_students = df.copy()
        remaining_departments = df["Department"].unique().tolist()
        available_departments = remaining_departments.copy()
        remaining_students_dict = {
            dept: {"students": [], "subject_code": department_subject_map.get(dept)}
            for dept in remaining_departments
        }

        # Step 1: Assign students to initial rooms
        while len(remaining_students) > 0:
            room = {"room_number": room_number, "side_a": [], "side_b": []}
            side_a, side_b = [], []

            if available_departments:
                dept_a = available_departments[0]
                dept_a_students = remaining_students[remaining_students["Department"] == dept_a]
                subject_code_a = department_subject_map[dept_a]

                if len(dept_a_students) >= 12:
                    while len(side_a) < 24 and not dept_a_students.empty:
                        student = dept_a_students.iloc[0]
                        side_a.append((student['Department'], student['Roll Number'], subject_code_a))
                        remaining_students = remaining_students[remaining_students["Roll Number"] != student["Roll Number"]]
                        dept_a_students = dept_a_students[dept_a_students["Roll Number"] != student["Roll Number"]]

                    if dept_a_students.empty:
                        available_departments.remove(dept_a)
                else:
                    remaining_students_dict[dept_a]["students"].extend(
                        [(student['Department'], student['Roll Number']) for _, student in dept_a_students.iterrows()]
                    )
                    remaining_students = remaining_students[remaining_students["Department"] != dept_a]
                    available_departments.remove(dept_a)
                    continue

            if available_departments:
                for dept_b in available_departments:
                    if department_subject_map[dept_b] != subject_code_a:
                        dept_b_students = remaining_students[remaining_students["Department"] == dept_b]

                        if len(dept_b_students) >= 12:
                            while len(side_b) < 24 and not dept_b_students.empty:
                                student = dept_b_students.iloc[0]
                                side_b.append((student['Department'], student['Roll Number'], department_subject_map[dept_b]))
                                remaining_students = remaining_students[remaining_students["Roll Number"] != student["Roll Number"]]
                                dept_b_students = dept_b_students[dept_b_students["Roll Number"] != student["Roll Number"]]

                            if dept_b_students.empty:
                                available_departments.remove(dept_b)
                            break
                        else:
                            remaining_students_dict[dept_b]["students"].extend(
                                [(student['Department'], student['Roll Number']) for _, student in dept_b_students.iterrows()]
                            )
                            remaining_students = remaining_students[remaining_students["Department"] != dept_b]
                            available_departments.remove(dept_b)

            while len(side_a) < 24:
                side_a.append(("---", "---", None))
            while len(side_b) < 24:
                side_b.append(("---", "---", None))

            room["side_a"] = side_a
            room["side_b"] = side_b
            rooms.append(room)
            room_number += 1

        # Step 2: Assign Unassigned Students to New Rooms
        unassigned_students_list = []
        for dept, data in remaining_students_dict.items():
            for student in data["students"]:
                unassigned_students_list.append((dept, student[1], data["subject_code"]))

        index = 0
        total_unassigned = len(unassigned_students_list)

        while index < total_unassigned:
            room = {"room_number": room_number, "side_a": [], "side_b": []}
            side_a, side_b = [], []

            # Fill Side A first
            while len(side_a) < 24 and index < total_unassigned:
                side_a.append(unassigned_students_list[index])
                index += 1

            # Fill Side B ensuring no same subject codes in the same bench
            for i in range(24):
                if index >= total_unassigned:
                    break
                current_student = unassigned_students_list[index]

                # Ensure subject codes are different for Side A and Side B at the same bench
                if side_a[i][2] is None or side_a[i][2] != current_student[2]:
                    side_b.append(current_student)
                    index += 1
                else:
                    side_b.append(("---", "---", None))  # Keep it empty if subject codes match

            while len(side_a) < 24:
                side_a.append(("---", "---", None))
            while len(side_b) < 24:
                side_b.append(("---", "---", None))

            room["side_a"] = side_a
            room["side_b"] = side_b
            rooms.append(room)
            room_number += 1

        # Step 3: Generate Attendance Sheets
        output_dir = os.path.join(OUTPUT_DIR, "attendance_sheets")
        os.makedirs(output_dir, exist_ok=True)

        for room in rooms:
            room_file_path = os.path.join(output_dir, f"room_{room['room_number']}_attendance.xlsx")
            with pd.ExcelWriter(room_file_path) as writer:
                for department in remaining_departments:
                    side_a_data = [student[1] for student in room["side_a"] if student[0] == department]
                    side_b_data = [student[1] for student in room["side_b"] if student[0] == department]

                    dept_data = {
                        "Department": [department] * (len(side_a_data) + len(side_b_data)),
                        "Roll Number": side_a_data + side_b_data,
                        "Side": ["Side A"] * len(side_a_data) + ["Side B"] * len(side_b_data)
                    }

                    dept_df = pd.DataFrame(dept_data)
                    sheet_name = f"{department}_Room{room['room_number']}"

                    dept_df.to_excel(writer, index=False, sheet_name=sheet_name)

            print(f"Attendance sheet for Room {room['room_number']} generated at {room_file_path}.")

        # Zip the files
        zip_filename = os.path.join(output_dir, "attendance_sheets.zip")
        with zipfile.ZipFile(zip_filename, 'w') as zipf:
            for room in rooms:
                room_file_path = os.path.join(output_dir, f"room_{room['room_number']}_attendance.xlsx")
                zipf.write(room_file_path, os.path.basename(room_file_path))

        print(f"All attendance sheets have been zipped: {zip_filename}")

        return send_from_directory(output_dir, "attendance_sheets.zip", as_attachment=True)

    except Exception as e:
        print(f"Error during attendance sheet generation: {str(e)}")
        return "An error occurred while generating the attendance sheets. Please try again."


if __name__ == "__main__":
    app.run(debug=True)