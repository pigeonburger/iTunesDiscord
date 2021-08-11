import pypresence, win32com.client, time, sys
from tkinter import PhotoImage, Label, Tk, CENTER, Button

def exitWarn():
    win = Tk()

    # Apple Music PNG logo in Base64
    icon = "iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAIAAAD8GO2jAAAACXBIWXMAAAABAAAAAQBPJcTWAAAErklEQVR4nLXV+U8UZxgH8Pdv6WG1HHKJci249wnLcrjsclpAoXjUmiYmja3VX2rS1MQQGxHUUkCtFLF41PiLoqxypFV6iNVqS2M9UFhkj3nfmX36zLyFTt2xoU2YfLK7PO+T7zM77LxDWOX2JUWYb9uSIsy7ZUkRum7z/yaWb2Zlm2jJ2yK+vqSH4Np/U9JCi5slV6PobAh5mmL+bdC0c67yHVbaotlPcP6/YJ4mJBZuEBz11P4WhmIWNH8Q3NcBPWeeff8TPHwCj4ITHd3YEHM38X51AqHFGzWJ7r/Nbv4QdrfO9J2DwLfhJ09BYBLIhwhA8S0YvXP8FHXUgbuR96tzCJ4dYkUvgsIGdMvbBAeOCTFRCYQY/ONgCpiJ3O7uFaw12B+fRpirURM469FoZQucC2CKABqHwOt8gLka+1mhQpVDmHO9JrDVoVF/E5wZpHEDpPlXQblEE10nBaMf7HXxOYRhVQtYatFYxUb4+jJVLnds/hJF50Jwc2L2wiVx3+HHOz+Bu388PPpl2OAXbRo5hNpqNYG5Bl33NsLAZaDi1K+T0H8x+FknbProuX8rmKppQcWM3j9rrIKbd6cPnwjpfcxay6lziGip1gRGWcDbABcC0HtxuKwecssRW+uTDJWiqUoyVwfttbP44xn/5VnH8fBar2YOoaZKTaD3oUB5PXwTgM/7R9w1vMJXmVEWtFUh/AZP24+FdOW8iNQ5hBr8mqCgAgVK1sP5AHSelgcoFb4q6mVBs3/WUokDpg71hHPLeBGpc4hY4NMEOi8a8tTJA472j7iqeEXdEzT6Zk1+eUBbdzi7VDOHsPx1mkBXBnmlQ8W1cG4Ijpwacfjliq6Mr4qKGUMFgvGfpw59EVrjFrVyCM0r0wQ5JehKUbU84HDfsN3HK3xVzJUFC8qndaVw4/bUwc5wZhFW4nMIyynVBFkedBWvzNkh6Ogbtngh24P4qpTuEtOcDwxeybcF7j2ePNgVzSiMZWvkECHbrSm2uhAN2rx4o0H7V9fN63hFzHJHVzknt++CnoGpyd8hEoVgZLy9M5pqxVWa5UbqHMLWFC0QFBghZLog0QwJpkueGhj8DtpOXjOVQ7oL3ckrhq17JOWmxn0igm/Pwj+2dQorLbDKpU7jCMWx81im7Hm6E3+OTz9uhRNn5XPEqNauEb0HUhxo0FGBdwZuHtGFzW46NH6gPZpkggwnT1BnEpbhUjgQpDpiydZrG7bB2C2qbMUS3/H3d43lF0OyHZJs8kU7e5UPiPLVB9NX331fSjJDqp3nzGfKiJhmXyCl2cMJht/OX4Qw5WfHlG0OPj0yllcIyTZ0L8sJjTsERvm2GpEk2NP6g64IUmxIncYRmmJFTIEf5t7UP75yDQfwh8lfT5gtuyZW22LJFiSttAoJxke1W2H3ftix91FFM/6JRZ4Tj9CVlgX4j2VvrB3etRceBkMAIeWh9WT0hpDhDKVYxWQL4p3icr3wej7CD+qEeIQlmRZIibL76Vaoe0/qPy/0DuC539e5xRUFkGhQdy4eoYnGF4gJxsiyfPpKDoq8loPp8T2LR+gK/QJBQeO8rL4YhC3PX1KELdMtKcJezV1SfwLOiQH4H52DOwAAAABJRU5ErkJggg=="

    win.geometry("450x100")
    win.title("iTunes Discord Presence")
    win.tk.call('wm','iconphoto',win._w,PhotoImage(data=icon))
    Label(win, text= "Could not create Discord Rich Presence. Please open Discord and try again.", font= ('Helvetica 9 bold')).place(relx=.5, rely=.3, anchor=CENTER)
    Button(win, text= "Ok", background= "white", foreground= "blue", font= ('Helvetica 13 bold'), command=sys.exit).place(relx=.45, rely=.5)
    win.mainloop()
    sys.exit()

try:
    client = "874807932800348233"

    o = win32com.client.Dispatch("iTunes.Application")
    DiscordRPC = pypresence.Presence(client, pipe=0)
    DiscordRPC.connect()

    while not o.CurrentTrack:
        DiscordRPC.update(state="Idling", large_image="amlogo", large_text="Apple Music")
        time.sleep(1)
    
    track = o.CurrentTrack.Name

    artist = o.CurrentTrack.Artist
    state = o.PlayerState

    pos = o.PlayerPosition

    starttime = int(time.time()) - o.PlayerPosition
    endtime = int(time.time()) + (o.CurrentTrack.Duration - o.PlayerPosition)

    DiscordRPC.update(state=artist, details=track, start=starttime, end=endtime,large_image="amlogo",large_text="Listening to Apple Music",small_image="play",small_text="Playing")

    while True:
        if track == o.CurrentTrack.Name:
            starttime =   int(time.time()) - o.PlayerPosition
            endtime = int(time.time()) + (o.CurrentTrack.Duration - o.PlayerPosition)

            if pos == o.PlayerPosition:
                small_image = "pause"
                small_text = "Paused"
            else:
                small_image = "play"
                small_text = "Playing"

            DiscordRPC.update(state=artist, details=track, start=starttime, end=endtime, large_image="amlogo", large_text="Listening to Apple Music", small_image=small_image, small_text=small_text)

            pos = o.PlayerPosition
            time.sleep(2)
            pass

        track = o.CurrentTrack.Name
        artist = o.CurrentTrack.Artist
        state = o.PlayerState

        starttime =   int(time.time()) - o.PlayerPosition
        endtime = int(time.time()) + (o.CurrentTrack.Duration - o.PlayerPosition)

        if pos == o.PlayerPosition:
            small_image = "pause"
            small_text = "Paused"
        else:
            small_image = "play"
            small_text = "Playing"
        
        DiscordRPC.update(state=artist, details=track, start=starttime, end=endtime, large_image="amlogo", large_text="Listening to Apple Music", small_image=small_image, small_text=small_text)

        pos = o.PlayerPosition
        time.sleep(2)

except pypresence.exceptions.InvalidID: 
    exitWarn()
except pypresence.exceptions.InvalidPipe:
    exitWarn()
except win32com.client.pywintypes.com_error:
    exit()
