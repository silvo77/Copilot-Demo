import subprocess
import os



class ChapterManager:

    def export_chapters_metadata(self, video_file: str, metadata_filename: str = "FFMETADATAFILE.txt"):
        """
        Export ffmpeg chapters metadata file from sessions with chapter timings.
        """
        def get_video_duration(video_file):
            ffmpeg_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ffmpeg', 'ffprobe.exe')
            cmd = [ffmpeg_path, "-v", "error", "-show_entries", "format=duration", "-of", "default=noprint_wrappers=1:nokey=1", video_file]
            try:
                output = subprocess.check_output(cmd, stderr=subprocess.STDOUT).decode().strip()
                return float(output)
            except Exception as ex:
                print(f"Error getting video duration: {ex}")
                return None

        def extract_metadata(video_file):
            ffmpeg_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ffmpeg', 'ffmpeg.exe')
            cmd = [ffmpeg_path, "-i", video_file, "-f", "ffmetadata", "-"]
            result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
            if result.returncode != 0:
                print("Error extracting metadata:", result.stderr)
                return ""
            return result.stdout

        # Filter sessions with valid chapter_start_time and chapter_end_time
        chapters = [s for s in self.sessions if getattr(s, 'chapter', False) and s.chapter_start_time is not None and s.chapter_end_time is not None]
        if not chapters:
            print("No chapters with valid timings found.")
            return

        chapter_text = ""
        video_duration = get_video_duration(video_file)
        for i, s in enumerate(chapters):
            start = int(s.chapter_start_time * 1000)
            # For last chapter, end at video duration
            if i < len(chapters) - 1:
                end = int(chapters[i].chapter_end_time)
            else:
                if video_duration is None:
                    end = int(video_duration * 1000)
                else:
                    end = int(s.chapter_end_time)
            title = s.title.replace('\n', ' ').replace('\r', ' ')
            chapter_text += f"""
[CHAPTER]
TIMEBASE=1/1000
START={start}
END={end}
title={title}
"""

        # Extract original metadata and prepend it
        metadata = extract_metadata(video_file)
        with open(metadata_filename, "w", encoding="utf-8") as f:
            f.write(metadata)
            f.write(chapter_text)
        print(f"Chapters metadata written to {metadata_filename}")

    def add_chapters_to_video_file(self, video_file: str, metadata_filename: str, output_file: str = None):
        """
        Use ffmpeg to mux chapters metadata into a new video file.
        """
        ffmpeg_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ffmpeg', 'ffmpeg.exe')
        if output_file is None:
            output_file = f"{os.path.splitext(video_file)[0]}_chapters.mp4"
        cmd = [ffmpeg_path, "-i", video_file, "-i", metadata_filename, "-map_metadata", "1", "-codec", "copy", output_file]
        print(f"Adding chapters to video...")
        print(' '.join(cmd))
        subprocess.run(cmd, check=True)
        print(f"Chapters added to: {output_file}")