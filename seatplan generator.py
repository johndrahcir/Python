import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
from PIL import Image, ImageTk
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas as pdfcanvas
import os

# Define constants
IMAGE_WIDTH = 140
IMAGE_HEIGHT = 140
MAX_IMAGES = 25
GRID_COLS = 5
MARGIN = 30

def open_folder(image_frame):
    folder_path = filedialog.askdirectory(title="Select Folder")
    if folder_path:
        image_files = [file for file in os.listdir(folder_path) if file.endswith(('jpg', 'jpeg', 'png', 'gif'))]
        if not image_files:
            messagebox.showwarning("Warning", "No image files found in the selected folder.")
            return

        if len(image_files) != MAX_IMAGES:
            messagebox.showwarning("Warning", f"Expected 25 images, found {len(image_files)}.")
            return

        num_images = len(image_frame.grid_slaves())
        if num_images + len(image_files) > MAX_IMAGES:
            messagebox.showwarning("Warning", f"Maximum number of images reached ({MAX_IMAGES})")
            return

        # Sort the image files based on numerical names
        image_files.sort(key=lambda x: int(x.split('.')[0]))

        # Reverse the list to display in the desired order
        image_files.reverse()

        num_label = MAX_IMAGES
        for image_file in image_files:
            image_path = os.path.join(folder_path, image_file)
            image = Image.open(image_path)
            image = image.resize((IMAGE_WIDTH, IMAGE_HEIGHT), resample=Image.LANCZOS)  # Use Image.LANCZOS
            photo_image = ImageTk.PhotoImage(image)
            photo_image.filename = image_path

            # Calculate the row and column positions in reverse order
            row, col = calculate_grid_position(num_images)

            # Create a label with the image
            new_label = tk.Label(image_frame, image=photo_image)
            new_label.image = photo_image

            # Add number label on top-left corner of the image
            num_label_text = tk.Label(image_frame, text=str(num_label), font=("Arial", 12), bg="white")
            num_label_text.grid(row=row, column=col, sticky="nw", padx=2, pady=2)  # Set sticky to northwest corner

            new_label.grid(row=row, column=col)

            num_images += 1
            num_label -= 1




def calculate_grid_position(num_images):
    row = num_images // GRID_COLS
    if row % 2 == 0:
        col = GRID_COLS - 1 - num_images % GRID_COLS
    else:
        col = num_images % GRID_COLS
    return row, col

def start_seat_plan():
    # Hide the main menu window
    root.withdraw()

    if hasattr(root, 'seat_plan_window') and root.seat_plan_window.winfo_exists():
        messagebox.showinfo("Info", "A seat plan is already in progress.")
        return

    seat_plan_name = simpledialog.askstring("Seat Plan", "Enter the name of the seat plan:")
    if seat_plan_name:
        root.seat_plan_window = tk.Toplevel()
        root.seat_plan_window.title(f"Seat Plan - {seat_plan_name}")

        canvas = tk.Canvas(root.seat_plan_window, bg="white")
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = tk.Scrollbar(root.seat_plan_window, orient=tk.VERTICAL, command=canvas.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        canvas.configure(yscrollcommand=scrollbar.set)

        global image_frame
        image_frame = tk.Frame(canvas, bg="white")
        canvas.create_window((0, 0), window=image_frame, anchor="nw")

        canvas.bind('<Configure>', lambda event, canvas=canvas: canvas.configure(scrollregion=canvas.bbox("all")))

        button_open = tk.Button(root.seat_plan_window, text="Select Images Folder", command=lambda: open_folder(image_frame))
        button_open.pack(pady=10)

        # Add a "Back to Main Menu" button
        button_back = tk.Button(root.seat_plan_window, text="Back to Main Menu", command=back_to_main_menu)
        button_back.pack(pady=10)

        # Add a "Save as PDF" button
        button_save_pdf = tk.Button(root.seat_plan_window, text="Save as PDF", command=lambda: save_as_pdf(image_frame, seat_plan_name))
        button_save_pdf.pack(pady=10)

def save_as_pdf(image_frame, seat_plan_name):
    save_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
    if save_path:
        c = pdfcanvas.Canvas(save_path, pagesize=letter)
        border_width = 1  # Border width in points

        # Calculate the maximum available width and height for each image
        max_image_width = (letter[0] - (GRID_COLS + 1) * MARGIN) / GRID_COLS
        max_image_height = (letter[1] - (GRID_COLS + 1) * MARGIN) / GRID_COLS

        # Draw the seat plan name on the PDF
        c.drawString(MARGIN, letter[1] - MARGIN, f"Seat Plan: {seat_plan_name}")

        # Draw "BACK" in the upper-middle part of the page
        front_text = "BACK"
        front_text_width = c.stringWidth(front_text)
        x_front = (letter[0] - front_text_width) / 2
        y_front = letter[1] - MARGIN
        c.drawString(x_front, y_front, front_text)

         # Draw "FRONT" in the lower-middle part of the page
        front_text = "FRONT"
        front_text_width = c.stringWidth(front_text)
        x_front = (letter[0] - front_text_width) / 2
        y_front = MARGIN - 10  # Adjust to the bottom of the page
        c.drawString(x_front, y_front, front_text)

        num_label = MAX_IMAGES
        for label in image_frame.winfo_children():
            if isinstance(label, tk.Label) and hasattr(label, 'image') and label.image:
                x_offset = MARGIN + label.grid_info()['column'] * (max_image_width + MARGIN)
                y_offset = MARGIN + label.grid_info()['row'] * (max_image_height + MARGIN)
                image_path = label.image.filename
                image = Image.open(image_path)
                image_width, image_height = image.size

                # Calculate the scale factor to fit the image within the available space
                scale_factor_width = max_image_width / image_width
                scale_factor_height = max_image_height / image_height
                scale_factor = min(scale_factor_width, scale_factor_height)

                # Resize the image proportionally
                new_width = int(image_width * scale_factor)
                new_height = int(image_height * scale_factor)
                image = image.resize((new_width, new_height), resample=Image.LANCZOS)

                # Adjust offsets to include margin
                x_offset += (max_image_width - new_width) / 2
                y_offset += (max_image_height - new_height) / 2

                # Draw the image onto the PDF
                c.drawImage(image_path, x_offset, letter[1] - y_offset - new_height, width=new_width, height=new_height, mask='auto')

                # Add number label on top-left corner of the image in PDF
                c.drawString(x_offset + 5, letter[1] - y_offset - new_height - -100, str(num_label))

                num_label -= 1

        c.save()
        messagebox.showinfo("Info", f"Seat plan saved as PDF at:\n{save_path}")





def back_to_main_menu():
    # Destroy the seat plan window
    root.seat_plan_window.destroy()
    # Show the main menu window
    root.deiconify()

# Create the main application window
root = tk.Tk()
root.title("Main Menu")

button_start = tk.Button(root, text="Create a Seat Plan", command=start_seat_plan)
button_start.pack(pady=10)

button_exit = tk.Button(root, text="Exit", command=root.destroy)
button_exit.pack(pady=10)

root.mainloop()
