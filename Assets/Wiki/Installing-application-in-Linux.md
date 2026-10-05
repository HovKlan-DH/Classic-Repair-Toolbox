[Wiki Home](Home)

Run CRT on Linux, and get it into your application menu.

---

The Linux download is one file, `Classic-Repair-Toolbox.AppImage`, from the [releases page](https://github.com/HovKlan-DH/Classic-Repair-Toolbox/releases). It runs directly from wherever you have downloaded it - it does not install the program anywhere else.

Before the first start, make it executable - either in your file manager (in the file's properties, allow it to run as a program) or in a terminal:

```
chmod +x Classic-Repair-Toolbox.AppImage
./Classic-Repair-Toolbox.AppImage
```

The first start copies the hardware data (close to 1 GB) out of the AppImage into `~/.local/share/Classic-Repair-Toolbox/Data`, so it takes a while. Your settings, log, workbooks and drafts live beside it in `~/.local/share/Classic-Repair-Toolbox/`. To keep the data somewhere else, see [Command-line parameters](Commandline-parameters).

If you want to have the CRT application and icon available in your desktop's application menu, then you can install it with an application manager like e.g. **Gear Lever**. Just open **Gear Lever** and drag the CRT file in to it, and afterwards you will be able to access it nice and easily:

<img width="902" height="578" alt="image" src="https://github.com/user-attachments/assets/13edb9d5-8b61-4259-bcc7-e0986d88ed51" />
