from flask import Flask, render_template, request, session, redirect, url_for, send_from_directory
import pandas as pd
import zipfile
import os
import uuid
import re
import tempfile
import shutil
from flask import Flask, render_template, send_file
from io import BytesIO
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet

app = Flask(__name__)
app.secret_key = "supersecretkey12345"  # Mandatory for session handling
LATEST_ATTENDANCE_DIR = None
count=0
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


def seating_logic():
        session_id = session.get("session_id")
        department_subject_map = session.get("department_subject_map")

        if not (session_id and department_subject_map):
            return "Error: Missing data. Please submit the form first."

        temp_file_path = os.path.join(TEMP_DIR, f"{session_id}.csv")
        if not os.path.exists(temp_file_path):
            return "Error: Data file not found. Please submit the form again."
        df = pd.read_csv(temp_file_path)
        df.sort_values(by=["Department", "Roll Number"], inplace=True)
        department_counts = df["Department"].value_counts().to_dict()
        total_students = df.shape[0]
        print(total_students)
        # Debug print to verify the counts
        print("\n📊 Initial Student Count Per Department:")
        for dept, count in department_counts.items():
            print(f"{dept}: {count} students")

        rooms = []
        special_rooms=[]
        room_number = 1
        session["room_number"]=room_number
        print("rgrfgg",room_number)
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
        session["room_number"]=room_number
  

        def find_empty_side_for_dept(dept, subject_code):
            """Find an available room where all students of a department can be placed on one side."""
            for room in rooms:
                if all(x[0] == "---" for x in room["side_a"]):  # Check if Side A is empty
                    if not any(student[2] == subject_code for student in room["side_b"] if student[2] is not None):
                        return room, "side_a"

                if all(x[0] == "---" for x in room["side_b"]):  # Check if Side B is empty
                    if not any(student[2] == subject_code for student in room["side_a"] if student[2] is not None):
                        return room, "side_b"

            return None, None

        def assign_students_to_room(room, side, students):
            """Assign students to a room side, ensuring exactly 24 slots are filled."""
            side_students = []

            while len(side_students) < 24 and students:
                side_students.append(students.pop(0))

            while len(side_students) < 24:  # Fill empty slots
                side_students.append(("---", "---", None))

            room[side] = side_students


        def create_new_special_room():
            """Create and return a new special room."""
            room_number=session.get("room_number")

        
            new_room = {"room_number": room_number, "side_a": [], "side_b": []}
            room_number += 1 
            
            return new_room  # Return the room before adding it to special_rooms


        def assign_unassigned_students():
            """Assign unassigned students while avoiding subject conflicts in special rooms."""
            unassigned_students_list = []
            
            # Collect all unassigned students
            for dept, data in remaining_students_dict.items():
                for student in data["students"]:
                    unassigned_students_list.append((dept, student[1], data["subject_code"]))

            print(f"\n🔎 Total Unassigned Students: {len(unassigned_students_list)}")
            room_number=session.get("room_number")
            # Group by department
            department_groups = {}
            for student in unassigned_students_list:
                dept, roll_number, subject_code = student
                if dept not in department_groups:
                    department_groups[dept] = []
                department_groups[dept].append(student)

            for dept, students in department_groups.items():
                print(f"\n🔹 Assigning Students from Department: {dept}")

                subject_code = students[0][2]  

                # First, check if a normal room is available
                room, side_to_fill = find_empty_side_for_dept(dept, subject_code)

                if room and side_to_fill:
                    print(f"✅ Placing {dept} students in Room {room['room_number']} on {side_to_fill}")
                    assign_students_to_room(room, side_to_fill, students)

                else:
                    # No normal room found → Try placing in an **existing special room sequentially**
                    assigned = False

                    for special_room in special_rooms:
                        # **Fill Side A first (sequentially)**
                        if len(special_room["side_a"]) < 24:
                            print(f"✅ Placing {dept} students in Special Room {special_room['room_number']} on Side A sequentially")
                            
                            while students and len(special_room["side_a"]) < 24:
                                special_room["side_a"].append(students.pop(0))
                            
                            assigned = True

                        # **Once Side A is full, move to Side B (check subject conflicts)**
                        if students and len(special_room["side_b"]) < 24:
                            # Ensure subject code conflict does not happen
                            if not any(student[2] == subject_code for student in special_room["side_a"] if student[2] is not None):
                                print(f"✅ Placing {dept} students in Special Room {special_room['room_number']} on Side B sequentially")

                                while students and len(special_room["side_b"]) < 24:
                                    special_room["side_b"].append(students.pop(0))

                                assigned = True

                        if assigned:
                            break  # Stop checking once students are placed

                    if not assigned:
                        # **Create new special room only if necessary**
                        print(f"⚠ No available room found. Creating new special room for {dept}.")
                        new_special_room = create_new_special_room()
                        
                        while students and len(new_special_room["side_a"]) < 24:
                            new_special_room["side_a"].append(students.pop(0))

                        special_rooms.append(new_special_room)
                        print(f"🆕 Special Room {new_special_room['room_number']} created and filled sequentially.")

        assign_unassigned_students()
        rooms.extend(special_rooms)
        return rooms,remaining_departments


@app.route("/seating_plan")
def seating_plan():
   
    try:
        # Retrieve session metadata
        rooms,remaining_departments=seating_logic()
        
        return render_template(
        "result.html",rooms=rooms)
    
    except Exception as e:
        return f"An error occurred: {str(e)}"
    


@app.route("/generate_attendance_sheets")
def generate_attendance_sheets():
    try:
        rooms,remaining_departments=seating_logic()
        # delete_contents(output_dir)
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

            # print(f"Attendance sheet for Room {room['room_number']} generated at {room_file_path}.")
            LATEST_ATTENDANCE_DIR = output_dir


        zip_filename = os.path.join(output_dir, "attendance_sheets.zip")
        count=0
        with zipfile.ZipFile(zip_filename, 'w') as zipf:
            for room in rooms:
                count+=1
                room_file_path = os.path.join(output_dir, f"room_{room['room_number']}_attendance.xlsx")
                zipf.write(room_file_path, os.path.basename(room_file_path))

        print(f"All attendance sheets have been zipped: {zip_filename}")
        print("COUNT:",count)
        LATEST_ATTENDANCE_DIR = output_dir
        session['LATEST_ATTENDANCE_DIR']=output_dir
        session['count']=count
        
        return render_template("download_redirect.html")

    except Exception as e:
        print(f"Error during attendance sheet generation: {str(e)}")
        return "An error occurred while generating the attendance sheets. Please try again."

@app.route("/download_attendance")
def download_attendance():
    output_dir = os.path.join(OUTPUT_DIR, "attendance_sheets")
    return send_from_directory(output_dir, "attendance_sheets.zip", as_attachment=True)

def get_seating_plan():
    LATEST_ATTENDANCE_DIR=session.get('LATEST_ATTENDANCE_DIR')
    count=session.get('count')

    if not LATEST_ATTENDANCE_DIR:
        print("Error: No latest attendance sheet directory found.")
        return {}

    directory = LATEST_ATTENDANCE_DIR
    seating_plan = {}

    for filename in os.listdir(directory):
        l=filename.split("_")
        t=count
        if filename.endswith(".xlsx") and filename.startswith("room_") and t:
            if int(l[1])<=count:
                t-=1
                print(filename)
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
