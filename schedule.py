import os
import cv2
import gspread
import requests
import subprocess
import numpy as np
from datetime import datetime
from pdfCropMargins import crop
from screeninfo import get_monitors
from pdf2image import convert_from_path
from oauth2client.service_account import ServiceAccountCredentials

key = 'SHEET_KEY'
cred = ServiceAccountCredentials.from_json_keyfile_name(
	'credentials.json',
	[
		'https://www.googleapis.com/auth/spreadsheets',
		'https://www.googleapis.com/auth/drive'
	]
)

auth = gspread.authorize(cred)
sched = auth.open_by_key(key)
export_url = f'https://www.docs.google.com/spreadsheets/export?format=pdf&id={key}'
bearer = cred.create_delegated('').get_access_token().access_token

pdf = requests.get(
	export_url, 
	headers={
		'Authorization': f'Bearer {bearer}'
	}
)

fn = f'schedule_{datetime.now()}'
pdf_name, crop_name, img_name = fn+'.pdf', fn+'_cropped.pdf', fn+'.png'
with open(pdf_name, 'wb') as fb:
	fb.write(pdf.content)

crop(['-p', '0', pdf_name])
img = convert_from_path(crop_name, 1500)[0]
img.save(img_name, 'PNG')

cur_time = datetime.now()
may_third = cur_time.replace(month=5, day=3)
five_am = cur_time.replace(hour=5, minute=0, second=0, microsecond=0)
midnight = cur_time.replace(hour=23, minute=59, second=59, microsecond=0)

n1 = (cur_time - may_third).days + 1
n2 = (cur_time - five_am).seconds // 60

n2 = n2 // 30 + 1
if not (five_am <= cur_time <= midnight):
	n1 -= 1

x = int(4+222*(n1+0.5))
y = int(52*(n2+0.5))

crop_w, crop_h = 222, 800
img = cv2.imread(img_name)
h, w, _ = img.shape
crop_range = [
	max(0, y-crop_h//2),
	min(y+crop_h//2, h),
	max(0, x-crop_w//2),
	min(x+crop_w//2, w)
]
crop_img = img[crop_range[0]:crop_range[1], crop_range[2]:crop_range[3]]
crop_time = img[crop_range[0]:crop_range[1], 6:228]
sched_img = np.hstack([crop_time, crop_img])

monitor = get_monitors()[0]
monitor_h, monitor_w = monitor.height, monitor.width
monitor_ratio = monitor_w / monitor_h

wall_img = cv2.imread('images/wall.jpeg')
h, w, _ = wall_img.shape
new_h = w // monitor_ratio
wall_img = wall_img[int(h//2-new_h//2):int(h//2+new_h//2)]
h, w, _ = wall_img.shape

scale = 0.4
sched_h, sched_w, _ = sched_img.shape
new_sched_h = int(h * scale) // 2 * 2
new_sched_w = int(sched_w * new_sched_h / sched_h) // 2 * 2
sched_img = cv2.resize(sched_img, (new_sched_w, new_sched_h))

crop_range = [
	int(h//2-new_sched_h//2),
	int(h//2+new_sched_h//2),
	int(w//2-new_sched_w//2),
	int(w//2+new_sched_w//2)
]
wall_img[crop_range[0]:crop_range[1], crop_range[2]:crop_range[3]] = sched_img

final_img_pth = 'images/save.jpeg'
cv2.imwrite(final_img_pth, wall_img)

def get_script(script):
	out = f"""
osascript<<END
	{script}
END
"""
	return out

def change_desktop(cur_desktop):
	change_desktop = f"""
tell application "System Events"
set targetDesktop to item {cur_desktop} of desktops
set targetDesktop's picture to "{os.path.abspath(final_img_pth)}"
end tell
"""
	return change_desktop

def get_desktop_n():
	num_desktop = """
tell application "System Events"
get the number of desktops
end tell
"""
	desktop_n = subprocess.check_output(get_script(num_desktop), shell=True)
	desktop_n = int(desktop_n.decode('utf-8'))
	return desktop_n

refresh = """killall Dock"""

desktop_n = get_desktop_n()
for cur_desktop in range(1, desktop_n+1):
	change = get_script(change_desktop(cur_desktop))
	subprocess.check_call(change, shell=True)

subprocess.check_call(refresh, shell=True)

os.remove(pdf_name)
os.remove(crop_name)
os.remove(img_name)
