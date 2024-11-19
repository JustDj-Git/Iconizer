# 🎨  Iconizer 🚀

The script consists of two parts. The first part, **`pull`**, involves extracting 32x32 or 16x16 icons from executable files (`.exe`) and saving them in different formats. The second part, **`apply`**, customizes folder icons by locating `.exe` or `.ico` files and creating a `desktop.ini` file with the required settings to display an icon on a folder, making both tools useful for developers and users who need efficient icon management solutions.

## 🛠️  Icon Extraction (`pull`):
  - Extracts icons from executable files (.exe).
  - Saves icons in multiple formats: ICO, BMP, PNG, JPG.
  - Allows selection of files through a dialog if not provided.
  - Supports extraction of small (16x16) or standard (32x32) icons.
  - Can specify the index of the icon to extract.
  - Provides logging functionality to track extraction processes.

## 🛠️  Folder Customization (`apply`):
  - Customizes folder icons using found .exe or .ico files.
  - Creates a desktop.ini file with required settings for folder icons.
  - Supports specifying multiple folder paths for icon application.
  - Offers priority settings for icon selection
  - Allows filtering of specific folder names to ignore during icon application.
  - Provides an option to apply icons only to specified folders, excluding subfolders.
  - Supports custom icon search rules using a hashtable.
  - Includes an option to remove icons from folders and restart Explorer.

## 📚 How to use `pull` and `apply`

### 🖱️ With GUI

1.   Open PowerShell (not CMD). Right-click on the Windows start menu and find PowerShell (or Terminal), or press `Win + S` and type Powershell.
2.   Copy and paste the code below and press enter for invoking **`pull`** function only

```powershell
irm icon.scripts.wiki | iex; pull
```

or for invoking **`apply`** function only

```powershell
irm icon.scripts.wiki | iex; apply
```

or for both

```powershell
irm icon.scripts.wiki | iex; pull; apply
```

---

## 🔧 Parameters and Switches for `pull`
 📝 The script accepts the following parameters and switches:

  - **`-directory` or `-d`**: An array of paths to `.exe` files or directories. If not provided, the user can select a file via a dialog (ex. `-file "C:\path\to\", "C:\path\to\test.exe"`).
  - **`-format` or `-f`**: Specifies the format for saving the icon (default is ICO). Acceptable values are `ico`, `bmp`, `png`, and `jpg` (ex. `-format ico`).
  - **`-index` or `-i`**: The index of the icon to extract from the executable (ex. `-index 1`).
  - **`-depth` or `-dep`**: The Depth parameter specifies the number of subdirectory levels to include in the recursion (ex. `-dep 2`).
  - **`-small` or `-s`**: A switch to indicate whether to extract the small 16x16 icon (default behavior is to extract the 32x32 icon).
  - **`-log` or `-l`**: Enables logging, writing to the specified file. Accepts a full path and file name (ex. `-log 'C:\task_log.txt'`).

### ⚙️ Examples with `pull`

1. **Use .ico image format, size 32x32 and path 'C:\location\text.exe'**

	```powershell
	irm ico.scripts.wiki | iex; pull -d 'C:\location\text.exe' -f 'ico'
	```

2. **Use .png image format, 16x16 icon with index 1 and multiple folder paths**  

	```powershell
	irm ico.scripts.wiki | iex; pull -d 'C:\location', 'C:\location2' -f 'png' -index '1' -s
	```

### Notes
- If **`directory`** is not specified, a system FileDialog will open.
- If **`format`** is not specified, `ico` will be used as the default format.
- If **`index`** is not specified, index `0` will be used by default.

---
## 🔧 Parameters and Switches for `apply`
📝 The script accepts the following parameters and switches:

- **`-directory` or `-d`**: Specify the folder paths in the format `'FIRST_PATH'` or `@('FIRST_PATH', 'SECOND_PATH')`, or `'FIRST_PATH', 'SECOND_PATH'`. If not set, the system FolderDialog will open.
- **`-priority` or `-p`**: Sets the priority for `.ico` files:
	- `icon`: An icon.ico file takes priority over .exe files and other .ico files.
	- `any`: Any .ico file takes priority over .exe files.
	- `folder`: A .ico file with the same name as the folder takes priority over .exe files and other .ico files.
- **`-filter` or `-f`**: Specify the names of folders to ignore in the format `'FIRST_NAME'` or `'FIRST_NAME', 'SECOND_NAME'`, or `@('FIRST_NAME', 'SECOND_NAME')`.
- **`-single` or `-s`**: When used, this switch applies the icon exclusively to the specified folder, excluding any subfolders.
- **`-dependencies` or `-dep`**: Use a hashtable to define custom rules for icon searches, e.g., `@{"Inno Setup 5" = "Compil32.exe"; "DaVinci Resolve" = "Resolve.exe"}`.
- **`-remove` or `-rm`**: This switch removes the folder icons and restarts Explorer.

---

### ⚙️ Examples with `apply`

1. **Passing two folders: `C:\test` and `D:\test2`**

	```powershell
	irm icon.scripts.wiki | iex; apply -d 'C:\test', 'D:\test2'
	```

2. **Passing two folders:  `C:\test` and `D:\test2` with filtered folders 'Windows Kits' and 'OneCommander'**

	```powershell
	irm icon.scripts.wiki | iex; apply -d 'C:\test', 'D:\test2' -f 'Windows Kits', 'OneCommander'
	```

3. **Passing folder `C:\test` with prioritizing `icon.ico` icons over `.exe` and folder name**

	```powershell
	irm icon.scripts.wiki | iex; apply -d 'C:\test' -p icon
	```

4. **Passing folder`C:\test` and assign an icon only to the specified folder without recursion**

	```powershell
	irm icon.scripts.wiki | iex; apply -d 'C:\test' -s
	```

5. **Remove icons from folders `C:\test` and `D:\test2`, don't remove from `Windows Kits` folder and restart Explorer**

	```powershell
	irm icon.scripts.wiki | iex; apply -d 'C:\test', 'D:\test2' -f 'Windows Kits' -rm
	```
---
### ⚙️ Examples for both `pull` and `apply`
1. **Passing two folders: `C:\test` and `D:\test2`, with `2` folder depth, icon format is `.ico`. Than  apply the icon to the folders inside `D:\test2`**

	```powershell
	irm icon.scripts.wiki | iex; pull -d 'C:\test', 'D:\test2' -dep '2' -f ico; apply -d 'D:\test2'
	```