This folder should also include the `AppData.zip`.

# Contents of `AppData.zip`

- Collected From: `%AppData%\Caliper\Maptitude 2020\`
- Extracted To: `C:\Users\Default\AppData\Roaming\Caliper\Maptitude 2020\`

I really only cared about the following two files:

- `checkupdate.arr`: gets created if *Edit* your *Preferences* and uncheck the *System Startup* setting for *Check for Updates on Startup*.
- `softwarereg.arr`: gets created if check the "Don't show this message again" when prompted for *Online Registration*.

Unfortunately, the `softwarereg.arr` doesn't appear to affect the application's first run.
I went ahead and took the entire AppData folder because something in it allowed the `checkupdate.arr` to work.
Without the rest of the AppData folder, `checkupdate.arr` wouldn't have worked either.

I would really prefer a registry setting under HKLM somewhere for both of these.
Hopefully, Caliper can improve this.
