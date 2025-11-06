import pyautogui
import time
import os

# Enable fail-safe
pyautogui.FAILSAFE = True

# Step 1: Open Outlook via Windows Search
pyautogui.press('win')
time.sleep(1)
pyautogui.write('Outlook')
time.sleep(1)
pyautogui.press('enter')
time.sleep(8)  # Wait for Outlook to load

# Step 2: Open "New Email" (Ctrl + N)
pyautogui.hotkey('ctrl', 'n')
time.sleep(3)

# Step 3: Type recipient
pyautogui.write("vishnupriya.r@sharpsyssoft.com")
pyautogui.press("tab")  # Move to Subject

# Step 4: Type subject
pyautogui.write("Asking for the status of the daily activity")
pyautogui.press("tab")  # Move to body

# Step 5: Type email body
email_body = [
    "Dear Vishnu priya,",
    "",
    "I hope this message finds you well. I would like to understand the status of your pending activities",
    "",
    "Please share it in a detailed way",
    "",
    "",
    "",
    "Thank you.",
    "",
    "Warm regards,",
    "Mathan Kumar R",
    "Functional trainee"
]

for line in email_body:
    pyautogui.write(line)
    pyautogui.press("enter")

time.sleep(1)

# Step 6: Attach file using keyboard shortcuts
# Shortcut: Alt + N → AF (Attach File)
pyautogui.hotkey('alt', 'n')
time.sleep(1)
pyautogui.press('a')
pyautogui.press('f')
time.sleep(2)

# Step 7: Type file path and press Enter
resume_path = r"C:\Users\HP\Desktop\AI learnings\Social Eagle\vs code\SocialEagle\tasktracking.xlsx"
pyautogui.write(resume_path)
pyautogui.press("enter")
time.sleep(3)

# Step 8: Save as draft (Ctrl + S)
pyautogui.hotkey('ctrl', 's')
time.sleep(1)

# Step 9: Close the window (Alt + F4)
pyautogui.hotkey('alt', 'f4')
time.sleep(1)

print("✅ Email drafted and saved successfully using keyboard shortcuts only.")
