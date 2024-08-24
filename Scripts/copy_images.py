import os
import shutil

def copy_images(source_dir, dest_dir):
    print("Running")
    # Create destination directory if it doesn't exist
    if not os.path.exists(dest_dir):
        print("Directory does not exist")
        os.makedirs(dest_dir)

    # Traverse the source directory recursively
    for root, dirs, files in os.walk(source_dir):
        for file in files:
            # Check if the file name contains a dash character
            if '-' not in file and 'bak' not in file:
                # Form the source and destination paths
                src_path = os.path.join(root, file)
                # Append -800px to file name
                new_filename = f"{os.path.splitext(file)[0]}_800px{os.path.splitext(file)[1]}"
                dest_path = os.path.join(dest_dir, new_filename)
                #dest_path = os.path.join(dest_dir, file)
                # Copy the file to the destination directory
                shutil.copy(src_path, dest_path)
                print(f"File '{file}' copied to '{dest_dir}'.")


source_dir = 'D:/PICTURES_WEB/images/'
dest_dir = 'D:/PICTURES_WEB/large/'
copy_images(source_dir, dest_dir)
