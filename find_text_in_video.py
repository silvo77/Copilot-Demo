import sys
import subprocess
import pytesseract
from PIL import Image, ImageEnhance, ImageFilter
import io
import os
import time
import argparse
import logging
from datetime import datetime

# Setup logging
def setup_logging(video_file: str) -> str:
    """Setup logging configuration"""
    # Create logs directory if it doesn't exist
    logs_dir = "logs"
    if not os.path.exists(logs_dir):
        os.makedirs(logs_dir)
    
    # Create log filename based on video filename and timestamp
    video_name = os.path.splitext(os.path.basename(video_file))[0]
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = os.path.join(logs_dir, f"{video_name}_{timestamp}.log")
    
    # Configure logging
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file),
            logging.StreamHandler()  # Also print to console
        ]
    )
    return log_file

def seconds_to_hms(seconds: float) -> str:
    """Convert seconds to HH:MM:SS format"""
    hours = int(seconds // 3600)
    minutes = int((seconds % 3600) // 60)
    seconds = seconds % 60
    return f"{hours:02d}:{minutes:02d}:{seconds:05.2f}"

def preprocess_image(image, crop_area=None):
    """
    Preprocess image for better OCR results:
    - Crop to region of interest (if specified)
    - Convert to grayscale
    - Enhance contrast
    - Sharpen
    - Resize (2x)
    
    Args:
        image: PIL Image object
        crop_area: Tuple of (left, top, right, bottom) in percentage of image size
    """
    if crop_area is not None:
        # Convert percentages to pixels
        width, height = image.size
        left = int(width * crop_area[0] / 100)
        top = int(height * crop_area[1] / 100)
        right = int(width * crop_area[2] / 100)
        bottom = int(height * crop_area[3] / 100)
        image = image.crop((left, top, right, bottom))
    
    gray = image.convert('L')
    enhancer = ImageEnhance.Contrast(gray)
    gray_enhanced = enhancer.enhance(2.0)
    sharpened = gray_enhanced.filter(ImageFilter.SHARPEN)
    width, height = sharpened.size
    resized = sharpened.resize((width*2, height*2), Image.LANCZOS)
    return resized

def find_text_in_video(video_file: str, start_time: float, duration: float, target_text: str, frame_rate: float = 1.0, crop_area=None, save_frames: bool = True) -> tuple[float, float, str]:
    """
    Search for text in a video file within a specified time window centered around start_time.
    
    Args:
        video_file: Path to the video file
        start_time: Center time point for the search window in seconds
        duration: Total duration of the search window in seconds (will be split before/after start_time)
        target_text: Text to search for
        frame_rate: Frames per second to extract (default: 1.0)
        crop_area: Tuple of (left, top, right, bottom) percentages to crop the image (default: None)
                  Each value should be between 0 and 100
        save_frames: Whether to save the frames where text is found (default: True)
    
    Returns:
        Tuple of (timestamp where text was found in seconds, elapsed processing time, path to saved frame)
        If text is not found, timestamp and frame path will be None
    """
    logging.info(f"Starting text search in video: {video_file}")
    logging.info(f"Search parameters: start_time={start_time}, duration={duration}, target_text='{target_text}', fps={frame_rate}")
    if crop_area:
        logging.info(f"Crop area: left={crop_area[0]}%, top={crop_area[1]}%, right={crop_area[2]}%, bottom={crop_area[3]}%")

    ffmpeg_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ffmpeg', 'ffmpeg.exe')
    if not os.path.exists(ffmpeg_path):
        ffmpeg_path = 'ffmpeg'  # Use system ffmpeg if local copy not found
        logging.info("Using system ffmpeg")
    else:
        logging.info("Using local ffmpeg")

    if crop_area:
        if not all(0 <= x <= 100 for x in crop_area) or len(crop_area) != 4:
            error_msg = "crop_area deve essere una tupla di 4 valori tra 0 e 100 (left, top, right, bottom)"
            logging.error(error_msg)
            raise ValueError(error_msg)
        if crop_area[0] >= crop_area[2] or crop_area[1] >= crop_area[3]:
            error_msg = "I valori right/bottom devono essere maggiori dei valori left/top"
            logging.error(error_msg)
            raise ValueError(error_msg)

    # Calculate the search window
    half_duration = duration / 2
    search_start = max(0, start_time - half_duration)  # Ensure we don't go below 0
    search_duration = duration
    
    # If we had to adjust the start time, extend the end time to maintain the total duration
    if search_start == 0 and start_time - half_duration < 0:
        # Add the "lost" time to the end of the search window
        lost_time = abs(start_time - half_duration)
        search_duration = duration + lost_time
        logging.info(f"Adjusted search window: extended end time by {lost_time:.2f}s due to negative start time")

    target_text = target_text.lower()
    cmd = [
        ffmpeg_path,
        '-ss', str(search_start),
        '-t', str(search_duration),
        '-i', video_file,
        '-vf', f'fps={frame_rate}',
        '-f', 'image2pipe',
        '-vcodec', 'png',
        '-hide_banner',
        '-loglevel', 'error',
        '-'
    ]

    area_info = f" (Area: L={crop_area[0]}%, T={crop_area[1]}%, R={crop_area[2]}%, B={crop_area[3]}%)" if crop_area else ""
    logging.info(f"Starting search from {seconds_to_hms(search_start)} to {seconds_to_hms(search_start + search_duration)}")
    logging.debug(f"FFmpeg command: {' '.join(cmd)}")
    print(f"Searching for '{target_text}' from {seconds_to_hms(search_start)} to {seconds_to_hms(search_start + search_duration)} at {frame_rate} fps{area_info}...")
    start_time_timer = time.time()
    
    try:
        process = subprocess.Popen(cmd, stdout=subprocess.PIPE, bufsize=10**8)
        frame_num = 0
        png_signature = b'\x89PNG\r\n\x1a\n'
        data = b""
        
        while True:
            chunk = process.stdout.read(8192)
            if not chunk:
                break
            
            data += chunk
            while True:
                start_idx = data.find(png_signature)
                if start_idx == -1:
                    break
                    
                next_idx = data.find(png_signature, start_idx + 8)
                if next_idx == -1:
                    break
                    
                png_data = data[start_idx:next_idx]
                data = data[next_idx:]
                
                try:
                    image = Image.open(io.BytesIO(png_data))
                    processed_image = preprocess_image(image, crop_area)
                    text = pytesseract.image_to_string(processed_image).lower()
                    timestamp = search_start + (frame_num / frame_rate)
                    
                    if frame_num % 10 == 0:  # Log progress every 10 frames
                        logging.info(f"Processing frame {frame_num} at {seconds_to_hms(timestamp)}")
                    
                    if target_text in text:
                        elapsed = time.time() - start_time_timer
                        logging.info(f"Text found in frame {frame_num} at {seconds_to_hms(timestamp)}")
                        
                        frame_filename = None
                        if save_frames:
                            # Save the frame with timestamp in the filename
                            base_name = os.path.splitext(os.path.basename(video_file))[0]
                            timestamp_str = seconds_to_hms(timestamp).replace(':', '_')
                            frame_filename = f"{base_name}_frame_{timestamp_str}.png"
                            
                            # Save both original and processed frames
                            image.save(frame_filename)
                            processed_filename = f"{base_name}_frame_{timestamp_str}_processed.png"
                            processed_image.save(processed_filename)
                            
                            logging.info(f"Saved original frame as: {frame_filename}")
                            logging.info(f"Saved processed frame as: {processed_filename}")
                        
                        logging.info(f"Total processing time: {seconds_to_hms(elapsed)}")
                        
                        print(f"Text found at {seconds_to_hms(timestamp)} (frame {frame_num})")
                        if save_frames:
                            print(f"Frame saved as: {frame_filename}")
                            print(f"Processed frame saved as: {processed_filename}")
                        print(f"Elapsed time: {seconds_to_hms(elapsed)}")
                        process.terminate()
                        return timestamp, elapsed, frame_filename
                        
                    frame_num += 1
                    
                except Exception as e:
                    error_msg = f"Error processing frame {frame_num}: {str(e)}"
                    logging.error(error_msg)
                    print(error_msg)
                    
        process.wait()
        
    except KeyboardInterrupt:
        interrupt_msg = "\nInterrupted by user"
        logging.warning(interrupt_msg)
        print(interrupt_msg)
        process.terminate()
        process.wait()
        
    except Exception as e:
        error_msg = f"Error: {str(e)}"
        logging.error(error_msg)
        print(error_msg)
        if process:
            process.terminate()
            process.wait()
            
    elapsed = time.time() - start_time_timer
    logging.info(f"Search completed. Text not found. Total time: {seconds_to_hms(elapsed)}")
    print(f"Text not found. Elapsed time: {seconds_to_hms(elapsed)}")
    return None, elapsed, None

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Search for text in a video file")
    parser.add_argument('video', type=str, help='Path to video file')
    parser.add_argument('text', type=str, help='Text to search for')
    parser.add_argument('--start', type=float, default=0.0, help='Start time in seconds (default: 0.0)')
    parser.add_argument('--duration', type=float, default=60.0, help='Duration to search in seconds (default: 60.0)')
    parser.add_argument('--fps', type=float, default=1.0, help='Frames per second to analyze (default: 1.0)')
    parser.add_argument('--left', type=float, help='Left crop boundary in percentage (0-100)')
    parser.add_argument('--top', type=float, help='Top crop boundary in percentage (0-100)')
    parser.add_argument('--right', type=float, help='Right crop boundary in percentage (0-100)')
    parser.add_argument('--bottom', type=float, help='Bottom crop boundary in percentage (0-100)')
    parser.add_argument('--no-save-frames', action='store_true',
                       help='Non salvare le immagini dei frame dove viene trovato il testo')

    args = parser.parse_args()

    # Verify video file exists
    if not os.path.exists(args.video):
        print(f"Error: Video file '{args.video}' not found")
        sys.exit(1)

    # Initialize logging
    log_file = setup_logging(args.video)
    logging.info("=== Starting new search session ===")
    logging.info(f"Log file: {log_file}")

    # Setup crop area if any boundary is specified
    crop_area = None
    if any(x is not None for x in [args.left, args.top, args.right, args.bottom]):
        if not all(x is not None for x in [args.left, args.top, args.right, args.bottom]):
            error_msg = "Error: When specifying crop area, all boundaries (--left, --top, --right, --bottom) must be provided"
            logging.error(error_msg)
            print(error_msg)
            sys.exit(1)
        crop_area = (args.left, args.top, args.right, args.bottom)

    found_time, elapsed, frame_path = find_text_in_video(
        args.video,
        args.start,
        args.duration,
        args.text,
        args.fps,
        crop_area,
        save_frames=not args.no_save_frames
    )

    if found_time is not None:
        logging.info("=== Search completed successfully ===")
        print(f"\nResults:")
        print(f"- Text found at: {seconds_to_hms(found_time)}")
        print(f"- Frame saved as: {frame_path}")
        print(f"- Processing time: {seconds_to_hms(elapsed)}")
        sys.exit(0)
    else:
        logging.info("=== Search completed without finding text ===")
        print(f"\nText not found in the specified time range")
        sys.exit(1)
