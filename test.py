import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from tkinter import ttk
from PIL import Image, ImageTk
import pillow_heif
import time

# Helper functions
def allowed_file(filename):
    ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'gif', 'tiff', 'pdf', 'doc', 'docx', 'heic', 'webp'}
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def open_image(filepath):
    ext = filepath.rsplit('.', 1)[1].lower()
    if ext == 'heic':
        pillow_heif.register_heif_opener()
        img = Image.open(filepath)
    else:
        img = Image.open(filepath)
    return img

# Function to resize images
def resize_image(img, width, height):
    return img.resize((width, height), Image.Resampling.LANCZOS)

# Function to convert and compress an image to meet the target size
def compress_image_within_size(img, save_path, format, target_size_kb):
    quality_low = 10  # Minimum quality
    quality_high = 95  # Maximum quality
    step_limit = 20  # Limit number of attempts
    step = 0

    # First, check if converting to JPEG will help reduce size more effectively
    if format != "JPEG":
        img = img.convert("RGB")  # Convert to RGB for JPEG
        format = "JPEG"
        save_path = save_path.replace(".png", ".jpg").replace(".webp", ".jpg")  # Ensure correct extension

    while step < step_limit:
        current_quality = (quality_low + quality_high) // 2
        img.save(save_path, format=format, optimize=True, quality=current_quality)
        file_size_kb = os.path.getsize(save_path) / 1024  # Get size in KB

        print(f"Step {step+1}: Quality={current_quality}, File Size={file_size_kb} KB, Target={target_size_kb} KB")

        if abs(file_size_kb - target_size_kb) <= 5:
            print("Target size achieved.")
            break
        elif file_size_kb > target_size_kb:
            quality_high = current_quality - 1
        else:
            quality_low = current_quality + 1

        # Stop early if no further compression is possible
        if quality_low >= quality_high or current_quality <= 10:
            break
        step += 1
        root.update_idletasks()
        time.sleep(0.1)

    # If the size is still larger, attempt to resize the image dimensions
    final_file_size_kb = os.path.getsize(save_path) / 1024
    if final_file_size_kb > target_size_kb:
        print("Warning: Could not compress to target size with quality adjustments. Attempting to resize dimensions.")
        # Resizing the image to reduce size further
        width, height = img.size
        resize_factor = 0.9
        while final_file_size_kb > target_size_kb and width > 100 and height > 100:  # Ensure it doesn't become too small
            width = int(width * resize_factor)
            height = int(height * resize_factor)
            img_resized = resize_image(img, width, height)
            img_resized.save(save_path, format=format, optimize=True, quality=quality_low)
            final_file_size_kb = os.path.getsize(save_path) / 1024
            print(f"Resizing: Width={width}, Height={height}, New File Size={final_file_size_kb} KB")

# Function to rename a single file
def rename_file(filepath, destination_folder, index, prefix):
    # Create a new file name with the given prefix and index
    file_extension = os.path.splitext(filepath)[1]  # Get the file extension
    new_filename = f"{prefix}{index + 1}{file_extension}"
    save_path = os.path.join(destination_folder, new_filename)
    
    # Copy the file to the destination folder with the new name
    shutil.copy(filepath, save_path)

# Function to process a single file (compress, rename, resize, or convert)
def process_file(filepath, destination_folder=None, index=0, prefix=None):
    if allowed_file(filepath):
        try:
            # Handle renaming
            if action_var.get() == "Rename":
                if not prefix:
                    prefix = "renamed_"  # Default prefix if not provided
                rename_file(filepath, destination_folder, index, prefix)
                messagebox.showinfo("Success", f"File renamed to {prefix}{index + 1}")
            else:
                # Open image if it's not a PDF
                img = open_image(filepath)

                # Resize option
                if action_var.get() == "Resize":
                    width = int(width_entry.get())
                    height = int(height_entry.get())
                    img = resize_image(img, width, height)

                # Compression option
                if action_var.get() == "Compress":
                    # Allow user to choose where to save the compressed file
                    save_path = filedialog.asksaveasfilename(defaultextension=f".{format_var.get()}",
                                                             filetypes=[(f"{format_var.get().upper()} files", f"*.{format_var.get()}")])
                    if save_path:
                        save_format = format_var.get().upper()
                        if save_format == "JPG":
                            save_format = "JPEG"

                        # Convert RGBA images to RGB before saving as JPEG
                        if save_format == "JPEG" and img.mode == "RGBA":
                            img = img.convert("RGB")

                        # Get target size in KB from the user
                        target_size_kb = simpledialog.askinteger("Target Size", "Enter target size in KB (e.g. 200 KB):")
                        if target_size_kb:
                            compress_image_within_size(img, save_path, save_format, target_size_kb)
                        else:
                            messagebox.showwarning("Cancelled", "Compression cancelled by the user.")
                    else:
                        messagebox.showwarning("Cancelled", "File save operation was cancelled.")
                else:
                    # Allow user to select save path for conversion
                    if processing_mode.get() == "Single File":
                        save_path = filedialog.asksaveasfilename(defaultextension=f".{format_var.get()}",
                                                                 filetypes=[(f"{format_var.get().upper()} files", f"*.{format_var.get()}")])
                    else:
                        # Folder operation: save to destination folder with modified name
                        save_path = os.path.join(destination_folder, f"{os.path.splitext(os.path.basename(filepath))[0]}_{index+1}.{format_var.get()}")

                    if save_path:
                        save_format = format_var.get().upper()
                        if save_format == "JPG":
                            save_format = "JPEG"

                        # Convert RGBA to RGB for JPEG format
                        if save_format == "JPEG" and img.mode == "RGBA":
                            img = img.convert("RGB")

                        # Save the converted file
                        img.save(save_path, format=save_format)
                    else:
                        messagebox.showwarning("Cancelled", "File save operation was cancelled.")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to process file: {e}")
    else:
        messagebox.showerror("Error", "The selected file format is not supported.")


# Function to compress PDF within size limit (in KB)
def compress_pdf_within_size(filepath, save_path, target_size_kb):
    try:
        reader = PdfReader(filepath)
        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)
        with open(save_path, "wb") as output_pdf:
            writer.write(output_pdf)
        # Check the size and compress by removing objects if needed
        while os.path.getsize(save_path) / 1024 > target_size_kb:
            writer.remove_blank_pages()
            with open(save_path, "wb") as output_pdf:
                writer.write(output_pdf)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to compress PDF: {e}")

# Function to process PDF files (including compression)
def process_pdf(filepath, destination_folder, index):
    try:
        # Allow user to choose where to save the compressed PDF
        save_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if save_path:
            # Get target size in KB from the user
            target_size_kb = simpledialog.askinteger("Target Size", "Enter target size in KB (e.g. 500 KB):")
            if target_size_kb:
                compress_pdf_within_size(filepath, save_path, target_size_kb)
            else:
                messagebox.showwarning("Cancelled", "Compression cancelled by the user.")
        else:
            messagebox.showwarning("Cancelled", "File save operation was cancelled.")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to process PDF: {e}")
# Function to process a folder (calls process_file for each file)
def process_folder(source_folder, destination_folder):
    files = [os.path.join(source_folder, file) for file in os.listdir(source_folder) if allowed_file(file)]
    if not files:
        messagebox.showerror("Error", "No valid files in the selected folder.")
        return

    prefix = None
    if action_var.get() == "Rename":
        prefix = simpledialog.askstring("Rename Keyword", "Enter keyword for renaming:", parent=root)
        if not prefix:
            messagebox.showerror("Error", "Rename keyword is required!")
            return
    
    total_files = len(files)
    progress["maximum"] = total_files  # Set the maximum value of the progress bar
    
    for idx, file in enumerate(files):
        process_file(file, destination_folder, idx, prefix)
        
        # Update the progress label and bar for each processed file
        progress_label.config(text=f"Progress: {idx + 1}/{total_files} files processed")
        progress_label.update_idletasks()  # Force the UI to refresh
        progress["value"] = idx + 1  # Update progress bar value
        progress.update_idletasks()  # Force the progress bar to refresh

    messagebox.showinfo("Success", f"All {total_files} files have been processed and saved to {destination_folder}")

# Tkinter UI with folder processing option
def browse_file():
    filepath = filedialog.askopenfilename()
    if filepath:
        selected_file_label.config(text=filepath)

def browse_source_folder():
    folderpath = filedialog.askdirectory()
    if folderpath:
        source_folder_label.config(text=folderpath)

def browse_destination_folder():
    folderpath = filedialog.askdirectory()
    if folderpath:
        destination_folder_label.config(text=folderpath)

def toggle_processing_mode(*args):
    if processing_mode.get() == "Single File":
        browse_file_button.config(state="normal")
        browse_source_folder_button.config(state="disabled")
        browse_destination_folder_button.config(state="disabled")
    else:
        browse_file_button.config(state="disabled")
        browse_source_folder_button.config(state="normal")
        browse_destination_folder_button.config(state="normal")

def toggle_action_options(*args):
    if action_var.get() == "Resize":
        width_label.grid()
        width_entry.grid()
        height_label.grid()
        height_entry.grid()
        format_label.grid_remove()
        format_menu.grid_remove()
    elif action_var.get() == "Convert" or action_var.get() == "Compress":
        format_label.grid()
        format_menu.grid()
        width_label.grid_remove()
        width_entry.grid_remove()
        height_label.grid_remove()
        height_entry.grid_remove()
    elif action_var.get() == "Rename":
        width_label.grid_remove()
        width_entry.grid_remove()
        height_label.grid_remove()
        height_entry.grid_remove()
        format_label.grid_remove()
        format_menu.grid_remove()

def start_processing():
    progress.start()
    if processing_mode.get() == "Single File":
        filepath = selected_file_label.cget("text")
        if os.path.isfile(filepath):
            process_file(filepath, destination_folder=destination_folder_label.cget("text"))
            messagebox.showinfo("Success", "Single file processed successfully!")
        else:
            messagebox.showerror("Error", "No valid file selected!")
    elif processing_mode.get() == "Folder Operation":
        source_folder = source_folder_label.cget("text")
        destination_folder = destination_folder_label.cget("text")
        if os.path.isdir(source_folder) and os.path.isdir(destination_folder):
            process_folder(source_folder, destination_folder)
        else:
            messagebox.showerror("Error", "Source or destination folder not selected!")
    progress.stop()


# Tkinter UI with folder processing option
def browse_file():
    filepath = filedialog.askopenfilename()
    if filepath:
        selected_file_label.config(text=filepath)

# Function to set up window icon
def set_window_icon(root):
    icon_path = os.path.join(os.getcwd(), 'app_icon.webp')
    if os.path.exists(icon_path):
        app_icon = ImageTk.PhotoImage(file=icon_path)
        root.iconphoto(True, app_icon)

# Tkinter setup
root = tk.Tk()
root.title("Document Processing Tool")
root.geometry("600x500")
root.configure(bg="#f0f0f0")

# Set the window icon
set_window_icon(root)

# Main Frame
main_frame = ttk.Frame(root, padding=(20, 20))
main_frame.pack(padx=10, pady=10, fill="both", expand=True)

# Logo
try:
    logo_img = Image.open("cropped-wrl-logo-new.png")
    logo_img = logo_img.resize((170, 60), Image.LANCZOS)
    logo_photo = ImageTk.PhotoImage(logo_img)

    logo_label = ttk.Label(main_frame, image=logo_photo)
    logo_label.image = logo_photo
    logo_label.grid(row=0, column=0, columnspan=3, pady=(0, 10))
except Exception as e:
    print("Logo not found or unable to load:", e)

# Instructions Label with functionality
instructions = ttk.Label(main_frame, text=(
    "Instructions:\n"
    "1. For single file processing, select a file using 'Browse File' and then press 'Start Processing'.\n"
    "2. For folder operations, select both a source and a destination folder.\n"
    "3. Ensure you choose the appropriate action (Rename, Convert, Compress, or Resize)."
), font=("Verdana", 9, "bold"))

instructions.grid(row=1, column=0, columnspan=3, pady=(0, 10))

# Processing mode selection (Single File or Folder)
processing_mode = tk.StringVar(value="Single File")
ttk.Radiobutton(main_frame, text="Single File Operation", variable=processing_mode, value="Single File").grid(row=2, column=0, padx=10, pady=5, sticky="w")
ttk.Radiobutton(main_frame, text="Folder Operation", variable=processing_mode, value="Folder Operation").grid(row=2, column=1, padx=10, pady=5, sticky="w")
processing_mode.trace("w", toggle_processing_mode)

# File selection (Single File)
selected_file_label = ttk.Label(main_frame, text="No file selected", width=50)
selected_file_label.grid(row=3, column=0, padx=10, pady=10, columnspan=2, sticky='w')
browse_file_button = ttk.Button(main_frame, text="Browse File", command=browse_file)
browse_file_button.grid(row=3, column=2, padx=10, pady=10)

# Source folder selection (For Folder Operation)
source_folder_label = ttk.Label(main_frame, text="No source folder selected", width=50)
source_folder_label.grid(row=4, column=0, padx=10, pady=10, columnspan=2, sticky='w')
browse_source_folder_button = ttk.Button(main_frame, text="Source Folder", command=browse_source_folder)
browse_source_folder_button.grid(row=4, column=2, padx=10, pady=10)

# Destination folder selection (For Folder Operation)
destination_folder_label = ttk.Label(main_frame, text="No destination folder selected", width=50)
destination_folder_label.grid(row=5, column=0, padx=10, pady=10, columnspan=2, sticky='w')
browse_destination_folder_button = ttk.Button(main_frame, text="Destination Folder", command=browse_destination_folder)
browse_destination_folder_button.grid(row=5, column=2, padx=10, pady=10)

# Action selection
action_var = tk.StringVar(value="Compress")

# Action section (Select Action)
action_label = ttk.Label(main_frame, text="Select Action:")
action_label.grid(row=6, column=0, padx=10, pady=5, sticky='e')

action_menu = ttk.OptionMenu(main_frame, action_var, "Compress", "Compress", "Resize", "Convert", "Rename")
action_menu.grid(row=6, column=1, padx=10, pady=5, sticky='w')
action_var.trace("w", toggle_action_options)

# Resize options (Width and Height) aligned below the "Select Action"
width_entry = ttk.Entry(main_frame)
width_label = ttk.Label(main_frame, text="Width:")
height_entry = ttk.Entry(main_frame)
height_label = ttk.Label(main_frame, text="Height:")

# Place Width and Height fields below the Select Action section
width_label.grid(row=7, column=0, padx=10, pady=5, sticky='e')
width_entry.grid(row=7, column=1, padx=10, pady=5, sticky='w')

height_label.grid(row=8, column=0, padx=10, pady=5, sticky='e')
height_entry.grid(row=8, column=1, padx=10, pady=5, sticky='w')

# Initially hide the resize options
width_label.grid_remove()
width_entry.grid_remove()
height_label.grid_remove()
height_entry.grid_remove()

# Format options for conversion
format_label = ttk.Label(main_frame, text="Convert to:")
format_label.grid(row=9, column=0, padx=10, pady=5, sticky='e')

format_var = tk.StringVar(value="png")
format_menu = ttk.OptionMenu(main_frame, format_var, "png", "png", "jpg", "gif", "tiff", "webp", "pdf")
format_menu.grid(row=9, column=1, padx=10, pady=5, sticky='w')

# Initially hide the convert options
format_label.grid_remove()
format_menu.grid_remove()

# Process button (aligned in the same row)
process_button = tk.Button(main_frame, text="Start Processing", command=start_processing, bg="#4CAF50", fg="black", font=("Verdana", 8, "bold"))
process_button.grid(row=10, column=0, padx=10, pady=30, columnspan=3, sticky='ew')

# Progress label to show real-time progress
progress_label = ttk.Label(main_frame, text="Progress: 0/0 files processed", font=("Verdana", 10))
progress_label.grid(row=11, column=0, columnspan=3, padx=10, pady=5)

# Progress bar
progress = ttk.Progressbar(main_frame, orient='horizontal', mode='determinate')
progress.grid(row=12, column=0, columnspan=3, padx=10, pady=5, sticky='ew')

# Configure grid weights
for i in range(3):
    main_frame.columnconfigure(i, weight=1)
main_frame.rowconfigure(12, weight=1)

# Start the application
root.mainloop()
