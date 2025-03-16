#!/usr/bin/env python3
"""
Things to Todoist Converter (AppleScript Version)
------------------------------------------------
This script uses AppleScript to extract data from Things and convert it to Todoist CSV format.
It preserves projects, areas, tasks, tags, due dates, and recurring task schedules.

Requirements:
- macOS (for AppleScript support)
- pandas (install with: pip install pandas) - Optional, falls back to csv module if not available

Usage:
1. Save this script as things_applescript_to_todoist.py
2. Make sure Things app is running
3. Run: python things_applescript_to_todoist.py <output_csv_path>
   Example: python things_applescript_to_todoist.py "todoist_import.csv"
"""

import sys
import subprocess
import json
import csv
from datetime import datetime
import argparse
import pickle
import os

# Try to import pandas, but provide a fallback if not available
try:
    import pandas as pd
    PANDAS_AVAILABLE = True
except ImportError:
    PANDAS_AVAILABLE = False
    print("Note: pandas is not installed. Using built-in csv module instead.")
    print("For better CSV handling, install pandas with: pip install pandas")

def run_applescript(script):
    """Run an AppleScript and return its output"""
    process = subprocess.Popen(['osascript', '-e', script], 
                               stdout=subprocess.PIPE, 
                               stderr=subprocess.PIPE)
    out, err = process.communicate()
    
    if process.returncode != 0:
        print(f"Error executing AppleScript: {err.decode('utf-8')}")
        return None
    
    return out.decode('utf-8').strip()

def get_areas():
    """Get all areas from Things via AppleScript"""
    script = '''
    tell application "Things3"
        set areaList to "["
        set allAreas to areas
        set areaCount to count of allAreas
        set currentIndex to 1
        
        repeat with theArea in allAreas
            set areaID to id of theArea
            set areaName to name of theArea
            
            -- Build JSON object manually
            set areaJSON to "{\\"id\\":\\"" & areaID & "\\",\\"name\\":\\"" & areaName & "\\"}"
            
            -- Add comma if not the last item
            if currentIndex < areaCount then
                set areaJSON to areaJSON & ","
            end if
            
            set areaList to areaList & areaJSON
            set currentIndex to currentIndex + 1
        end repeat
        
        set areaList to areaList & "]"
        return areaList
    end tell
    '''
    
    result = run_applescript(script)
    if result:
        try:
            return json.loads(result)
        except json.JSONDecodeError:
            print(f"Error parsing areas JSON: {result}")
            return []
    return []

def get_projects():
    """Get all projects from Things via AppleScript"""
    # Use a simpler approach with separate AppleScript calls
    
    # First, get all project IDs
    id_script = '''
    tell application "Things3"
        set projectIDs to id of projects
        return projectIDs
    end tell
    '''
    
    id_result = run_applescript(id_script)
    if not id_result:
        print("Error getting project IDs")
        return []
    
    # Parse the project IDs
    project_ids = []
    for item in id_result.strip().split(", "):
        item = item.strip()
        if item:  # Just check if the item is not empty
            project_ids.append(item)
    
    # Now get details for each project individually
    projects = []
    for project_id in project_ids:
        # Get project name
        name_script = f'''
        tell application "Things3"
            set theProject to project id "{project_id}"
            return name of theProject
        end tell
        '''
        name_result = run_applescript(name_script)
        
        # Get project notes
        notes_script = f'''
        tell application "Things3"
            set theProject to project id "{project_id}"
            return notes of theProject
        end tell
        '''
        notes_result = run_applescript(notes_script)
        
        # Get project status
        status_script = f'''
        tell application "Things3"
            set theProject to project id "{project_id}"
            set projectStatus to status of theProject
            if projectStatus is completed then
                return "completed"
            else
                return "active"
            end if
        end tell
        '''
        status_result = run_applescript(status_script)
        
        # Get project area
        area_script = f'''
        tell application "Things3"
            set theProject to project id "{project_id}"
            try
                set theArea to area of theProject
                if theArea is not missing value then
                    return id of theArea
                end if
            end try
            return "null"
        end tell
        '''
        area_result = run_applescript(area_script)
        
        # Build project data
        project_data = {
            "id": project_id,
            "name": name_result if name_result else "",
            "notes": notes_result if notes_result else "",
            "status": status_result if status_result else "active"
        }
        
        # Handle area
        if area_result and area_result != "null":
            project_data["area"] = area_result
        else:
            project_data["area"] = None
        
        projects.append(project_data)
        print(f"Processed project: {project_data['name']}")
    
    return projects

def get_to_dos():
    """Get all to-dos from Things via AppleScript"""
    # Use a simpler approach with separate AppleScript calls
    
    # First, collect all to-do IDs from inbox, projects, and areas
    todo_ids = []
    
    # Get to-dos from inbox
    inbox_script = '''
    tell application "Things3"
        set todoIDs to {}
        set inboxToDos to to dos of list "Inbox"
        repeat with theToDo in inboxToDos
            set end of todoIDs to id of theToDo
        end repeat
        return todoIDs
    end tell
    '''
    
    inbox_result = run_applescript(inbox_script)
    if inbox_result:
        for item in inbox_result.strip().split(", "):
            item = item.strip()
            if item.startswith("to do id "):
                todo_id = item[len("to do id "):]
                todo_ids.append((todo_id, None, None))  # (id, project_id, area_id)
    
    # Get to-dos from projects
    projects_script = '''
    tell application "Things3"
        set allProjects to projects
        set projectData to {}
        
        repeat with theProject in allProjects
            set projectID to id of theProject
            set projectToDos to to dos of theProject
            
            repeat with theToDo in projectToDos
                set todoID to id of theToDo
                set end of projectData to {todoID & "," & projectID}
            end repeat
        end repeat
        
        return projectData
    end tell
    '''
    
    projects_result = run_applescript(projects_script)
    if projects_result:
        for item in projects_result.strip().split(", "):
            item = item.strip().replace("{", "").replace("}", "")
            if "," in item:
                parts = item.split(",")
                todo_id = parts[0].strip()
                project_id = parts[1].strip()
                
                if todo_id.startswith("to do id "):
                    todo_id = todo_id[len("to do id "):]
                if project_id.startswith("project id "):
                    project_id = project_id[len("project id "):]
                
                todo_ids.append((todo_id, project_id, None))
    
    # Get to-dos from areas
    areas_script = '''
    tell application "Things3"
        set allAreas to areas
        set areaData to {}
        
        repeat with theArea in allAreas
            set areaID to id of theArea
            set areaToDos to to dos of theArea
            
            repeat with theToDo in areaToDos
                set todoID to id of theToDo
                set end of areaData to {todoID & "," & areaID}
            end repeat
        end repeat
        
        return areaData
    end tell
    '''
    
    areas_result = run_applescript(areas_script)
    if areas_result:
        for item in areas_result.strip().split(", "):
            item = item.strip().replace("{", "").replace("}", "")
            if "," in item:
                parts = item.split(",")
                todo_id = parts[0].strip()
                area_id = parts[1].strip()
                
                if todo_id.startswith("to do id "):
                    todo_id = todo_id[len("to do id "):]
                if area_id.startswith("area id "):
                    area_id = area_id[len("area id "):]
                
                todo_ids.append((todo_id, None, area_id))
    
    # Now get details for each to-do individually
    todos = []
    for todo_id, project_id, area_id in todo_ids:
        # Get basic to-do properties
        basic_script = f'''
        tell application "Things3"
            set theToDo to to do id "{todo_id}"
            set todoData to {{}}
            set end of todoData to name of theToDo
            set end of todoData to notes of theToDo
            set end of todoData to status of theToDo
            return todoData
        end tell
        '''
        
        basic_result = run_applescript(basic_script)
        if not basic_result:
            continue
        
        basic_parts = basic_result.strip().split(", ")
        if len(basic_parts) < 3:
            continue
        
        title = basic_parts[0].strip()
        notes = basic_parts[1].strip()
        status = basic_parts[2].strip()
        
        # Get due date
        due_date_script = f'''
        tell application "Things3"
            set theToDo to to do id "{todo_id}"
            try
                set dueDate to deadline of theToDo
                if dueDate is not missing value then
                    return dueDate as string
                end if
            end try
            return "null"
        end tell
        '''
        
        due_date_result = run_applescript(due_date_script)
        due_date = None if due_date_result == "null" else due_date_result
        
        # Get recurring info
        recurring_script = f'''
        tell application "Things3"
            set theToDo to to do id "{todo_id}"
            try
                if recurring of theToDo is true then
                    return recurrence of theToDo
                end if
            end try
            return "null"
        end tell
        '''
        
        recurring_result = run_applescript(recurring_script)
        recurring = None if recurring_result == "null" else recurring_result
        
        # Get tags
        tags_script = f'''
        tell application "Things3"
            set theToDo to to do id "{todo_id}"
            set tagList to {{}}
            try
                set todoTags to tags of theToDo
                repeat with theTag in todoTags
                    set end of tagList to name of theTag
                end repeat
            end try
            return tagList
        end tell
        '''
        
        tags_result = run_applescript(tags_script)
        tags = []
        if tags_result:
            for tag in tags_result.strip().split(", "):
                tags.append(tag.strip())
        
        # Get checklist items
        checklist_script = f'''
        tell application "Things3"
            set theToDo to to do id "{todo_id}"
            set checklistData to {{}}
            try
                set todoChecklist to checklistitems of theToDo
                repeat with checkItem in todoChecklist
                    set itemTitle to name of checkItem
                    set itemStatus to status of checkItem
                    set end of checklistData to {{itemTitle, itemStatus}}
                end repeat
            end try
            return checklistData
        end tell
        '''
        
        checklist_result = run_applescript(checklist_script)
        checklist = []
        if checklist_result and "{" in checklist_result:
            items = checklist_result.strip().split("}, {")
            for item in items:
                item = item.replace("{", "").replace("}", "")
                if "," in item:
                    parts = item.split(",", 1)
                    item_title = parts[0].strip()
                    item_status = parts[1].strip()
                    checklist.append({
                        "title": item_title,
                        "status": item_status
                    })
        
        # Build to-do data
        todo_data = {
            "id": todo_id,
            "title": title,
            "notes": notes,
            "status": status,
            "dueDate": due_date,
            "project": project_id,
            "area": area_id,
            "tags": tags,
            "recurring": recurring,
            "checklist": checklist
        }
        
        todos.append(todo_data)
        print(f"Processed to-do: {todo_data['title']}")
    
    return todos

def convert_things_recurrence_to_todoist(recurrence_str):
    """Convert Things recurrence format to Todoist format"""
    if not recurrence_str:
        return ""
    
    # Common patterns in Things recurrence strings
    recurrence = recurrence_str.lower()
    
    # Daily recurrence
    if "every day" in recurrence:
        return "every day"
    if "every weekday" in recurrence:
        return "every weekday"
    if "every" in recurrence and "days" in recurrence:
        # Try to extract number (e.g., "every 3 days")
        import re
        match = re.search(r'every\s+(\d+)\s+days', recurrence)
        if match:
            return f"every {match.group(1)} days"
    
    # Weekly recurrence
    weekdays = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]
    if "every week" in recurrence:
        for day in weekdays:
            if day in recurrence:
                return f"every week on {day.capitalize()}"
        return "every week"
    
    if "every" in recurrence and "weeks" in recurrence:
        match = re.search(r'every\s+(\d+)\s+weeks', recurrence)
        if match:
            interval = match.group(1)
            for day in weekdays:
                if day in recurrence:
                    return f"every {interval} weeks on {day.capitalize()}"
            return f"every {interval} weeks"
    
    # Monthly recurrence
    if "every month" in recurrence:
        # Check for specific day of month
        match = re.search(r'on the (\d+)(st|nd|rd|th)', recurrence)
        if match:
            return f"every month on the {match.group(1)}{match.group(2)}"
        return "every month"
    
    if "every" in recurrence and "months" in recurrence:
        match = re.search(r'every\s+(\d+)\s+months', recurrence)
        if match:
            interval = match.group(1)
            day_match = re.search(r'on the (\d+)(st|nd|rd|th)', recurrence)
            if day_match:
                return f"every {interval} months on the {day_match.group(1)}{day_match.group(2)}"
            return f"every {interval} months"
    
    # Yearly recurrence
    if "every year" in recurrence:
        return "every year"
    
    if "every" in recurrence and "years" in recurrence:
        match = re.search(r'every\s+(\d+)\s+years', recurrence)
        if match:
            return f"every {match.group(1)} years"
    
    # If no pattern matched, return a simplified version
    if "day" in recurrence:
        return "every day"
    if "week" in recurrence:
        return "every week"
    if "month" in recurrence:
        return "every month"
    if "year" in recurrence:
        return "every year"
    
    # Default fallback
    return "every day"

def convert_to_todoist_format(areas, projects, todos):
    """Convert Things data to Todoist format"""
    todoist_items = []
    
    # Create a mapping of Things area IDs to Todoist project positions
    area_to_todoist_position = {}
    for i, area in enumerate(areas):
        area_to_todoist_position[area['id']] = i

    # First add all areas (as top-level projects in Todoist)
    for area in areas:
        todoist_items.append({
            'TYPE': 'project',
            'CONTENT': area['name'],
            'PRIORITY': 1,
            'INDENT': 1,
            'AUTHOR': '',
            'RESPONSIBLE': '',
            'DATE': '',
            'DATE_LANG': 'en',
            'things_id': area['id']
        })
    
    # Then add all projects
    for project in projects:
        # Determine indent level (2 if it belongs to an area, 1 otherwise)
        indent = 2 if project['area'] and areas else 1
        
        # Skip completed projects if desired
        # if project['status'] == 'completed':
        #     continue
        
        # Mark completed projects
        project_name = project['name']
        if project['status'] == 'completed':
            project_name = f"✓ {project_name}"
        
        # Only use the mapping if we have areas and the project belongs to an area
        if areas and project['area'] and project['area'] in area_to_todoist_position:
            area_index = area_to_todoist_position[project['area']]
        else:
            area_index = 0
        
        todoist_items.append({
            'TYPE': 'project',
            'CONTENT': project_name,
            'PRIORITY': 1,
            'INDENT': indent,
            'AUTHOR': '',
            'RESPONSIBLE': '',
            'DATE': '',
            'DATE_LANG': 'en',
            'things_id': project['id']
        })
    
    # Then add all to-dos
    for todo in todos:
        # Determine if to-do is completed
        is_completed = (todo['status'] == 'completed')
        
        # Determine priority (Things doesn't have a built-in priority scale, using tags as a proxy)
        priority = 1  # Default normal priority
        if 'tags' in todo and todo['tags']:
            for tag in todo['tags']:
                tag_lower = tag.lower()
                if 'high' in tag_lower or 'urgent' in tag_lower or 'priority' in tag_lower:
                    priority = 4  # Highest in Todoist
                    break
        
        # Determine indent level and parent relationship
        indent = 1  # Default top-level
        if todo['project']:
            indent = 2  # Task in project
        elif todo['area']:
            indent = 2  # Task in area
        
        # Format due date and recurrence
        due_date = ''
        if todo['dueDate']:
            try:
                # Convert ISO format to YYYY-MM-DD if needed
                if 'T' in todo['dueDate']:
                    # It's an ISO format date
                    parsed_date = datetime.fromisoformat(todo['dueDate'].replace('Z', '+00:00'))
                    due_date = parsed_date.strftime('%Y-%m-%d')
                else:
                    # It's already a date string
                    due_date = todo['dueDate']
                
                # Add recurrence if present
                if todo['recurring']:
                    recurrence = convert_things_recurrence_to_todoist(todo['recurring'])
                    if recurrence:
                        due_date = f"{due_date} {recurrence}"
            except Exception as e:
                print(f"Error parsing date '{todo['dueDate']}': {e}")
        elif todo['recurring']:  # Recurring but no due date
            recurrence = convert_things_recurrence_to_todoist(todo['recurring'])
            if recurrence:
                due_date = recurrence
        
        # Prepare task content (title + notes if present)
        content = todo['title']
        if todo['notes']:
            content = f"{content}\n{todo['notes']}"
        
        # Add tags if present
        if 'tags' in todo and todo['tags']:
            tags_str = ' '.join([f"#{tag.replace(' ', '_')}" for tag in todo['tags']])
            content = f"{content} {tags_str}"
        
        # Add completion mark if task is completed
        if is_completed:
            content = f"✓ {content}"
        
        todoist_items.append({
            'TYPE': 'task',
            'CONTENT': content,
            'PRIORITY': priority,
            'INDENT': indent,
            'AUTHOR': '',
            'RESPONSIBLE': '',
            'DATE': due_date,
            'DATE_LANG': 'en'
        })
        
        # Add checklist items as subtasks
        if 'checklist' in todo and todo['checklist']:
            for item in todo['checklist']:
                item_content = item['title']
                if item['status'] == 'completed':
                    item_content = f"✓ {item_content}"
                
                todoist_items.append({
                    'TYPE': 'task',
                    'CONTENT': item_content,
                    'PRIORITY': 1,
                    'INDENT': indent + 1,  # One level deeper than parent
                    'AUTHOR': '',
                    'RESPONSIBLE': '',
                    'DATE': '',
                    'DATE_LANG': 'en'
                })
    
    # Remove temporary fields
    for item in todoist_items:
        if 'things_id' in item:
            item.pop('things_id')
    
    return todoist_items

def create_todoist_csv(items, output_path):
    """Create a Todoist-compatible CSV file"""
    print(f"Writing {len(items)} items to CSV: {output_path}")
    
    if PANDAS_AVAILABLE:
        # Use pandas if available (better handling of special characters and formatting)
        df = pd.DataFrame(items)
        df.to_csv(output_path, index=False)
    else:
        # Fallback to built-in csv module
        fieldnames = ['TYPE', 'CONTENT', 'PRIORITY', 'INDENT', 'AUTHOR', 'RESPONSIBLE', 'DATE', 'DATE_LANG']
        
        with open(output_path, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            for item in items:
                # Ensure only the expected fields are written
                filtered_item = {k: v for k, v in item.items() if k in fieldnames}
                writer.writerow(filtered_item)
    
    print("Conversion complete!")

def sanitize_text(text):
    """Sanitize text for CSV output"""
    if not text:
        return ""
    # Replace problematic characters
    return text.replace('\r', ' ').replace('\n\n', '\n')

def parse_args():
    parser = argparse.ArgumentParser(description='Convert Things data to Todoist CSV format')
    parser.add_argument('output', help='Output CSV file path')
    parser.add_argument('--skip-completed', action='store_true', help='Skip completed tasks')
    parser.add_argument('--areas-only', action='store_true', help='Only include areas')
    parser.add_argument('--projects-only', action='store_true', help='Only include projects')
    parser.add_argument('--todos-only', action='store_true', help='Only include to-dos')
    return parser.parse_args()

def save_cache(data, cache_file):
    with open(cache_file, 'wb') as f:
        pickle.dump(data, f)

def load_cache(cache_file):
    if os.path.exists(cache_file):
        with open(cache_file, 'rb') as f:
            return pickle.load(f)
    return None

def process_todos_in_batches(todo_ids, batch_size=50):
    """Process to-dos in batches to improve performance"""
    todos = []
    total_batches = (len(todo_ids) + batch_size - 1) // batch_size
    
    for i in range(0, len(todo_ids), batch_size):
        batch = todo_ids[i:i+batch_size]
        print(f"Processing batch {i//batch_size + 1}/{total_batches} ({len(batch)} to-dos)")
        # Process batch
        batch_todos = process_todo_batch(batch)
        todos.extend(batch_todos)
    
    return todos

def parse_things_date(date_str):
    """Parse a Things date string into a datetime object"""
    formats = [
        '%Y-%m-%d',
        '%Y-%m-%dT%H:%M:%S%z',
        '%Y-%m-%d %H:%M:%S',
        '%a, %d %b %Y %H:%M:%S %z'  # RFC 2822 format
    ]
    
    for fmt in formats:
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    
    return None

def print_help():
    """Print detailed help information"""
    print("""
Things to Todoist Converter
---------------------------
This script extracts data from Things 3 and converts it to Todoist CSV format.

Requirements:
- macOS (for AppleScript support)
- Things 3 app must be running
- pandas (optional, for better CSV handling)

Usage examples:
  python things_applescript_to_todoist.py output.csv
  python things_applescript_to_todoist.py output.csv --skip-completed
  python things_applescript_to_todoist.py output.csv --include-areas --include-projects --include-todos
    """)

def main():
    # Parse command-line arguments
    args = parse_args()
    output_csv = args.output
    
    print("Connecting to Things app via AppleScript...")
    
    # Check if Things is running
    script = 'tell application "System Events" to return exists process "Things3"'
    if run_applescript(script) != "true":
        print("Error: Things 3 is not running. Please open Things 3 and try again.")
        sys.exit(1)
    
    # Extract data from Things
    areas = []
    projects = []
    todos = []
    
    # Determine what to extract based on the flags
    include_areas = True
    include_projects = True
    include_todos = True
    
    # If any of the "only" flags are set, only include those items
    if args.areas_only or args.projects_only or args.todos_only:
        include_areas = args.areas_only
        include_projects = args.projects_only
        include_todos = args.todos_only
    
    if include_areas:
        print("Extracting areas...")
        areas = get_areas()
        print(f"Found {len(areas)} areas")
    
    if include_projects:
        print("Extracting projects...")
        projects = get_projects()
        print(f"Found {len(projects)} projects")
    
    if include_todos:
        print("Extracting to-dos...")
        todos = get_to_dos()
        print(f"Found {len(todos)} to-dos")
    
    # Filter out completed items if requested
    if args.skip_completed:
        projects = [p for p in projects if p['status'] != 'completed']
        todos = [t for t in todos if t['status'] != 'completed']
        print(f"After filtering: {len(projects)} projects, {len(todos)} to-dos")
    
    # Convert to Todoist format
    print("Converting to Todoist format...")
    todoist_items = convert_to_todoist_format(areas, projects, todos)
    
    # Create CSV
    create_todoist_csv(todoist_items, output_csv)
    
    print(f"\nImport instructions for Todoist:")
    print("1. Go to Todoist Settings > Import/Export")
    print("2. Choose 'Import from CSV file'")
    print(f"3. Select the generated file: {output_csv}")

if __name__ == "__main__":
    main()