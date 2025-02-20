from flask import Flask, render_template, request, session, redirect, url_for, send_from_directory
import pandas as pd
import zipfile
import os
import uuid
import re
from flask import Flask, render_template, send_file
from io import BytesIO
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet

app = Flask(__name__)
app.secret_key = "supersecretkey12345"  # Mandatory for session handling
LATEST_ATTENDANCE_DIR = None
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

        # Sort unassigned students department-wise
        unassigned_students_list.sort(key=lambda x: x[0])  # Sorting by department

        index = 0
        total_unassigned = len(unassigned_students_list)

        # Function to find available rooms with a fully unoccupied side
        def find_empty_side():
            for room in rooms:
                side_a_unplaced = all(x[0] == "---" for x in room["side_a"])
                side_b_unplaced = all(x[0] == "---" for x in room["side_b"])
                
                if side_a_unplaced:
                    return room, "side_a"
                elif side_b_unplaced:
                    return room, "side_b"
            
            return None, None

        while index < total_unassigned:
            room, side_to_fill = find_empty_side()

            if room and side_to_fill:
                # Fill the available side with one department's students
                dept_to_place = unassigned_students_list[index][0]  # Get department of the first unassigned student
                side_students = []

                while len(side_students) < 24 and index < total_unassigned and unassigned_students_list[index][0] == dept_to_place:
                    side_students.append(unassigned_students_list[index])
                    index += 1

                # Fill remaining spots with "---"
                while len(side_students) < 24:
                    side_students.append(("---", "---", None))

                room[side_to_fill] = side_students  # Update the room's side
            else:
                # If no empty rooms, create a new room
                room = {"room_number": room_number, "side_a": [], "side_b": []}
                side_a, side_b = [], []

                # Fill Side A first
                while len(side_a) < 24 and index < total_unassigned:
                    side_a.append(unassigned_students_list[index])
                    index += 1

                # Fill Side B ensuring different subject codes
                for i in range(24):
                    if index >= total_unassigned:
                        break
                    current_student = unassigned_students_list[index]

                    # Ensure subject codes are different for Side A and Side B at the same bench
                    if side_a[i][2] is None or (side_a[i][2] != current_student[2] and side_a[i][0] != current_student[0]):
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
        "result.html",rooms=rooms,
        buttons_visible=True,
        department_buttons_visible=session.get('department_buttons_visible', False)
    )
    
    except Exception as e:
        return f"An error occurred: {str(e)}"

@app.route("/generate_attendance_sheets")
def generate_attendance_sheets():
    global LATEST_ATTENDANCE_DIR
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

        # Sort unassigned students department-wise
        unassigned_students_list.sort(key=lambda x: x[0])  # Sorting by department

        index = 0
        total_unassigned = len(unassigned_students_list)

        # Function to find available rooms with a fully unoccupied side
        def find_empty_side():
            for room in rooms:
                side_a_unplaced = all(x[0] == "---" for x in room["side_a"])
                side_b_unplaced = all(x[0] == "---" for x in room["side_b"])
                
                if side_a_unplaced:
                    return room, "side_a"
                elif side_b_unplaced:
                    return room, "side_b"
            
            return None, None

        while index < total_unassigned:
            room, side_to_fill = find_empty_side()

            if room and side_to_fill:
                # Fill the available side with one department's students
                dept_to_place = unassigned_students_list[index][0]  # Get department of the first unassigned student
                side_students = []

                while len(side_students) < 24 and index < total_unassigned and unassigned_students_list[index][0] == dept_to_place:
                    side_students.append(unassigned_students_list[index])
                    index += 1

                # Fill remaining spots with "---"
                while len(side_students) < 24:
                    side_students.append(("---", "---", None))

                room[side_to_fill] = side_students  # Update the room's side
            else:
                # If no empty rooms, create a new room
                room = {"room_number": room_number, "side_a": [], "side_b": []}
                side_a, side_b = [], []

                # Fill Side A first
                while len(side_a) < 24 and index < total_unassigned:
                    side_a.append(unassigned_students_list[index])
                    index += 1

                # Fill Side B ensuring different subject codes
                for i in range(24):
                    if index >= total_unassigned:
                        break
                    current_student = unassigned_students_list[index]

                    # Ensure subject codes are different for Side A and Side B at the same bench
                    if side_a[i][2] is None or (side_a[i][2] != current_student[2] and side_a[i][0] != current_student[0]):
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


        # Step 3: Generate Attendance Sheets
        output_dir = os.path.join(OUTPUT_DIR, "attendance_sheets")
        os.makedirs(output_dir, exist_ok=True)
        for room in rooms:
            room_file_path = os.path.join(output_dir, f"room_{room['room_number']}_attendance.xlsx")

            with pd.ExcelWriter(room_file_path, engine='xlsxwriter') as writer:
                for department in remaining_departments:
                    # Filter students for this department in the current room
                    roll_numbers = [student[1] for student in room["side_a"] if student[0] == department] + \
                                [student[1] for student in room["side_b"] if student[0] == department]

                    if not roll_numbers:  
                        continue

                    
                    dept_df = pd.DataFrame({
                        "Roll Number": roll_numbers,
                        "Signature": [""] * len(roll_numbers)  
                    })

                    sheet_name = f"{department}_Room{room['room_number']}"[:31]  

                    dept_df.to_excel(writer, index=False, sheet_name=sheet_name, startrow=2)

                    workbook = writer.book
                    worksheet = writer.sheets[sheet_name]

                    title_format = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center'})
                    worksheet.merge_range('A1:B1', "DVR & Dr. HS MIC College of Technology", title_format)

                    dept_format = workbook.add_format({'bold': True, 'font_size': 12, 'align': 'center'})
                    worksheet.merge_range('A2:B2', f"Department Name: {department}", dept_format)

                    worksheet.set_column('A:A', 45) 
                    worksheet.set_column('B:B', 45)  
                    header_format = workbook.add_format({'bold': True, 'align': 'center', 'border': 1})
                    worksheet.write(2, 0, "Roll Number", header_format)  
                    worksheet.write(2, 1, "Signature", header_format) 
                    border_format = workbook.add_format({'border': 1, 'align': 'center'})

                    for row_idx in range(3, len(roll_numbers) + 3):  
                        worksheet.write(row_idx, 0, dept_df.iloc[row_idx - 3, 0], border_format)  
                        worksheet.write(row_idx, 1, "", border_format)  
                    for row_idx in range(2, len(roll_numbers) + 3):
                        worksheet.set_row(row_idx, 21.5)
                        
                    worksheet.set_margins(left=0.75, right=0, top=0.75, bottom=0.75)

            print(f"Attendance sheet for Room {room['room_number']} generated at {room_file_path}.")
            LATEST_ATTENDANCE_DIR = output_dir
            print(output_dir)

        zip_filename = os.path.join(output_dir, "attendance_sheets.zip")
        with zipfile.ZipFile(zip_filename, 'w') as zipf:
            for room in rooms:
                room_file_path = os.path.join(output_dir, f"room_{room['room_number']}_attendance.xlsx")
                zipf.write(room_file_path, os.path.basename(room_file_path))

        print(f"All attendance sheets have been zipped: {zip_filename}")
        LATEST_ATTENDANCE_DIR = output_dir
        print("generate attendance",LATEST_ATTENDANCE_DIR)
        session['department_buttons_visible'] = True 
        return render_template("download_redirect.html")

    except Exception as e:
        print(f"Error during attendance sheet generation: {str(e)}")
        return "An error occurred while generating the attendance sheets. Please try again."

@app.route("/download_attendance")
def download_attendance():
    output_dir = os.path.join(OUTPUT_DIR, "attendance_sheets")
    return send_from_directory(output_dir, "attendance_sheets.zip", as_attachment=True)

def get_seating_plan():
    global LATEST_ATTENDANCE_DIR
    print("get:"+LATEST_ATTENDANCE_DIR)
    if not LATEST_ATTENDANCE_DIR:
        print("Error: No latest attendance sheet directory found.")
        return {}

    directory = LATEST_ATTENDANCE_DIR
    seating_plan = {}

    for filename in os.listdir(directory):
        if filename.endswith(".xlsx") and filename.startswith("room_"):
            match = re.search(r"room_(\d+)", filename)
            if not match:
                continue  
            room_number = int(match.group(1))  
            file_path = os.path.join(directory, filename)

            try:
                xls = pd.ExcelFile(file_path)
                for sheet_name in xls.sheet_names:
                    df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
                    if df.shape[0] < 3:
                        continue

                    college_name = df.iloc[0, 0]
                    department = df.iloc[1, 0]

                    df = pd.read_excel(xls, sheet_name=sheet_name, skiprows=2)

                    if "Roll Number" not in df.columns:
                        continue

                    if not df.empty:
                        first_roll = df["Roll Number"].dropna().astype(str).min()
                        last_roll = df["Roll Number"].dropna().astype(str).max()

                        if department not in seating_plan:
                            seating_plan[department] = {"college": college_name, "rooms": []}

                        seating_plan[department]["rooms"].append({
                            "Room": room_number,
                            "First Roll": f"({first_roll})",
                            "Last Roll": f"({last_roll})"
                        })

            except Exception as e:
                print(f"Error processing file {filename}: {e}")

    return seating_plan

@app.route("/download_pdf")
def download_pdf():
    try:
        seating_plan = get_seating_plan()
        if not seating_plan:
            return "Error: Seating plan could not be generated.", 500

        buffer = BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=landscape(A4))
        elements = []
        styles = getSampleStyleSheet()

        for department, details in seating_plan.items():
            college_name = details["college"]
            rooms = details["rooms"]

            elements.append(Paragraph(f"<b>{college_name}</b>", styles["Title"]))
            elements.append(Spacer(1, 12))
            elements.append(Paragraph(f"<b>Department: {department}</b>", styles["Heading2"]))
            elements.append(Spacer(1, 12))

            data = [["Room Number", "FROM", "TO"]]
            for room in sorted(rooms, key=lambda x: x["Room"]):
                data.append([room["Room"], room["First Roll"], room["Last Roll"]])

            table = Table(data, colWidths=[100, 150, 150])
            table.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.grey),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("BOTTOMPADDING", (0, 0), (-1, 0), 10),
                ("GRID", (0, 0), (-1, -1), 1, colors.black),
            ]))

            elements.append(table)
            elements.append(PageBreak())

        if not elements:
            return "Error: No data available for the PDF.", 500

        doc.build(elements)
        buffer.seek(0)  # Reset buffer

        print("✅ PDF successfully generated!")

        return send_file(
            buffer,
            as_attachment=True,
            download_name="Seating_Plan.pdf",
            mimetype="application/pdf"
        )

    except Exception as e:
        print(f"❌ Error generating PDF: {e}")
        return "Error generating PDF.", 500



if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 10000)), debug=True)