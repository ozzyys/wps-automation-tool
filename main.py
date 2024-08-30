import pyautogui
import pyperclip
import pyfiglet
import time
import datetime

banner = pyfiglet.figlet_format("WPS AUTOMATION", font="slant")
print(banner)

screen_width, screen_height = pyautogui.size()
print(f"[*] Screen Size: {screen_width, screen_height}")

if screen_width != 1920:
    screen_width_multiplier = screen_width / 1920
else:
    screen_width_multiplier = 1

if screen_height != 1080:
    screen_height_multiplier = screen_height / 1080
else:
    screen_height_multiplier = 1

# while True:
#     current_mouse_x, current_mouse_y = pyautogui.position()
#     print(f"[*] Current Mouse Position: {current_mouse_x, current_mouse_y}")
#     time.sleep(0.2)

print(f"[*] Open WPS.")
pyautogui.press("win")
time.sleep(0.5)
pyautogui.typewrite("wps office")
pyautogui.press("enter")
time.sleep(1)

print(f"[*] Click new document.")
pyautogui.click(30 * screen_width_multiplier, 120 * screen_height_multiplier)
time.sleep(1)
pyautogui.click(300 * screen_width_multiplier, 228 * screen_height_multiplier)
time.sleep(2)

print(f"[*] Write your message.")
pyautogui.click(953, 604) # click the blank.
time.sleep(0.5)
pyperclip.copy("Dear Professors,\n\n"
               "Professor Constantine A. Balanis,the world-renowned authority in antenna theory and design, has honored us with a testimonial. His endorsement is a tremendous validation of our work and our commitment to advancing antenna technology.\n\n"
               "In addition, we’ve also received glowing testimonials from Prof. Dr. Erdem Topsakal, Prof. Dr. Rashaunda Henderson and Prof. Dr. Oscar Quevedo Teruel, each of whom has made significant contributions to our field. We are incredibly grateful for their support and recognition.\n\n"
               "We invite you to visit https://www.antenit.com/references/ to read these testimonials and learn more about how our block-based Anten’it products and innovations are making an impact in the antenna and microwave engineering communities.\n\n"
               "Thank you for being a part of our journey!\n\n"
               "Best regards,\n\n"
               "\n"
               "Ezgi Ozdemiroglu\n"
               "Business Development Engineer\n"
               "Antenom Antenna Technologies\n"
               "+905355603693\n"
               "www.antenit.com\n")
pyautogui.hotkey("ctrl", "v", interval=0.25) # copy text olduğundan bold çıkaramıyorum.
time.sleep(0.5)

print(f"[*] Click references.")
pyautogui.click(977 * screen_width_multiplier, 55 * screen_height_multiplier)
time.sleep(1)

print("[*] Click mail merge. Mail merge list open. So now we can see the 'Open Data Source' section.")
pyautogui.click(1502 * screen_width_multiplier, 117 * screen_height_multiplier)
time.sleep(1)

print("[*] Open Data Source.")
pyautogui.click(378 * screen_width_multiplier, 99 * screen_height_multiplier)
time.sleep(1)

print(f"[*] The file is being selected.")
pyautogui.write(f"Brevodan_cekilen_temiz_full_list(050724).xlsx")
pyautogui.press("down")
pyautogui.press("enter") # ilk enter dosyayı seçmek için
pyautogui.press("enter") # ikinci enter dosyayı onaylamak için
print(f"[*] The file has been selected.")
time.sleep(1)
pyautogui.press("enter") # üçüncü enter spawn olan box için
time.sleep(0.5)

# when program runs for second time, then references section already been clicked. after that "mail merge" section also
# been clicked.
# raise 1

starting_row = 12
ending_row = 150 + 12

while True:
    get_time = datetime.datetime.now()
    pyautogui.getWindowsWithTitle("WPS Office")[0].maximize() # WPS Office arka planda çalışmasına karşın ön plana al
    # üsttekinin işe yaramamasına karşın aşağıdaki kısımda tekrar wps'i aç. eski döküman korunmuş olacak.
    print(f"[*] Open WPS.")
    if starting_row == 0:
        pass
    else:
        pyautogui.press("win")
        time.sleep(0.1)
        pyautogui.typewrite("wps office")
        pyautogui.press("enter")

        time.sleep(1)

    print(f"[*] Click merge and send.")
    pyautogui.click(1481 * screen_width_multiplier, 96 * screen_height_multiplier)
    time.sleep(1)
    pyautogui.click(1472 * screen_width_multiplier, 141 * screen_height_multiplier) # send by email button.
    time.sleep(1)

    print(f"[*] Fill the blank lines.")
    pyautogui.click(937 * screen_width_multiplier, 442 * screen_height_multiplier)
    time.sleep(0.5)
    pyautogui.click(947 * screen_width_multiplier, 481 * screen_height_multiplier)
    time.sleep(0.5)

    if starting_row == 0:
        print(f"[*] Write the subject.")
        pyautogui.click(1027 * screen_width_multiplier, 467 * screen_height_multiplier)
        time.sleep(0.5)
        pyautogui.write("Testimonials") # subject line tek bir kez yazılmalı. ikincisinde, ilki hala duruyor oluyor.
        time.sleep(0.2)
    else:
        print(f"[*] The subject have already been entered.")

    print(f"[*] Select from section.")
    pyautogui.click(816, 585)
    time.sleep(0.2)

    print(f"[*] Click from section.")
    pyautogui.click(891, 585)
    time.sleep(0.2)
    pyautogui.hotkey("ctrl", "a")
    time.sleep(0.2)
    pyautogui.press("delete")
    time.sleep(0.2)
    pyautogui.write(f"{starting_row}")
    time.sleep(0.5)

    print(f"[*] Click to section.")
    pyautogui.click(986, 585)
    time.sleep(0.2)
    pyautogui.hotkey("ctrl", "a")
    time.sleep(0.2)
    pyautogui.press("delete")
    time.sleep(0.2)
    pyautogui.write(f"{ending_row}")
    time.sleep(0.5)

    print(f"[*] Send e-mail.")
    pyautogui.click(964 * screen_width_multiplier, 617 * screen_height_multiplier)
    time.sleep(0.5)
    print(f"[0] Mail has been sent via WPS-AUTOMATION tool. FROM: {starting_row}, TO: {ending_row}")

    starting_row += 1
    ending_row += 1
    time.sleep(3600) # gönderdikten sonra 1 saat bekle.
