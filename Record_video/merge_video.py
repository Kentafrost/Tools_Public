# merger all mp4 videos in a folder
from moviepy import VideoFileClip, concatenate_videoclips
import os
import logging
import tkinter as tk
from tkinter import Tk, filedialog

def merge_videos(input_folder):
    video_files = [f for f in os.listdir(input_folder) if f.endswith('.mp4')]
    video_files.sort()
    video_clips = []
    
    # select some mp4 files in the folder with button in explorer
    if not video_files:
        print("No mp4 files found in the selected folder.")
        return
    else:
        selected_videos = filedialog.askopenfilenames(
            title="Select MP4 files to merge",
            filetypes=[("MP4 files", "*.mp4")]
        )
    
    # all mp4 files into a list
    for video_file in selected_videos:
        video_path = os.path.join(input_folder, video_file)
        video_clips.append(VideoFileClip(video_path))
    
    if video_clips:
        try:
            final_clip = concatenate_videoclips(video_clips)
            
            output_file = os.path.join(input_folder, "merged_video.mp4")
            final_clip.write_videofile(output_file, codec='libx264')
            print(f"Merged video saved as {output_file}")
            
        except Exception as e:
            logging.error(f"Error merging videos: {e}")
            print(f"Error merging videos: {e}")
    else:
        print("No video files found to merge.")

    
if __name__ == "__main__":
    # Get the input folder from the user
    root = Tk()
    root.withdraw()  # Hide the root window
    input_folder = filedialog.askdirectory(title="Select Video Folder")
    
    if input_folder:
        merge_videos(input_folder)
    else:
        print("No folder selected. Exiting.")
        os._exit(0)