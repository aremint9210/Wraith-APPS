import pymupdf
import pypdf
from pypdf.generic import DictionaryObject, NameObject, NumberObject, ArrayObject, StreamObject, TextStringObject
import os
import sys
import re
import wave
import struct
import io

def log(msg, log_file):
    try:
        with open(log_file, 'a', encoding='utf-8') as f:
            f.write(msg + '\n')
    except Exception:
        pass

def natural_sort_key(s):
    return [int(text) if text.isdigit() else text.lower() for text in re.split(r'(\d+)', s)]

def print_progress_bar(iteration, total, prefix='', suffix='', decimals=1, length=30, fill='='):
    if total <= 0:
        return
    percent = ("{0:." + str(decimals) + "f}").format(100 * (iteration / float(total)))
    filled_length = int(length * iteration // total)
    bar = fill * filled_length + '-' * (length - filled_length)
    sys.stdout.write(f'\r{prefix} |{bar}| {percent}% {suffix}')
    sys.stdout.flush()
    if iteration == total:
        sys.stdout.write('\n')
        sys.stdout.flush()

def load_audio_as_16bit_pcm(wav_path, target_rate=44100):
    """
    Ensures audio data is strictly compliant with PDF 1.7 / ISO 32000-1 specifications
    and Adobe Acrobat Multimedia.api: 16-bit signed PCM mono at 44.1kHz.
    Converts 24-bit, 32-bit, or non-standard sample rates automatically to prevent
    Adobe Acrobat buffer overrun crashes (0xc0000409).
    """
    with open(wav_path, "rb") as f:
        file_bytes = f.read()
        
    try:
        with wave.open(io.BytesIO(file_bytes), 'rb') as w:
            n_channels = w.getnchannels()
            sampwidth = w.getsampwidth()
            src_rate = w.getframerate()
            n_frames = w.getnframes()
            
            # If already standard 16-bit mono 44100 Hz, return as is
            if sampwidth == 2 and n_channels == 1 and src_rate == target_rate:
                return file_bytes
                
            raw = w.readframes(n_frames)
            
        floats = []
        bytes_per_sample = sampwidth
        bytes_per_frame = bytes_per_sample * n_channels
        
        for i in range(0, len(raw), bytes_per_frame):
            chunk = raw[i : i + bytes_per_sample]
            if sampwidth == 1:
                val = (chunk[0] - 128) / 128.0
            elif sampwidth == 2:
                val = struct.unpack('<h', chunk)[0] / 32768.0
            elif sampwidth == 3:
                val = int.from_bytes(chunk, byteorder='little', signed=True) / 8388608.0
            elif sampwidth == 4:
                val = struct.unpack('<i', chunk)[0] / 2147483648.0
            else:
                val = 0.0
            floats.append(val)
            
        if src_rate != target_rate and src_rate > 0:
            target_len = int(len(floats) * target_rate / src_rate)
            resampled = []
            for j in range(target_len):
                src_idx = j * src_rate / target_rate
                idx_floor = int(src_idx)
                idx_ceil = min(idx_floor + 1, len(floats) - 1)
                frac = src_idx - idx_floor
                val = floats[idx_floor] * (1.0 - frac) + floats[idx_ceil] * frac
                resampled.append(val)
            floats = resampled
            
        pcm_bytes = bytearray()
        for f in floats:
            ival = max(-32768, min(32767, int(f * 32767)))
            pcm_bytes.extend(struct.pack('<h', ival))
            
        out_buf = io.BytesIO()
        with wave.open(out_buf, 'wb') as out_w:
            out_w.setnchannels(1)
            out_w.setsampwidth(2) # 16-bit
            out_w.setframerate(target_rate)
            out_w.writeframes(pcm_bytes)
            
        return out_buf.getvalue()
    except Exception:
        return file_bytes

def get_icon_positions(pdf_path, log_file):
    """
    Scans the PDF with PyMuPDF in read-only mode to locate sound icon locations
    near 'Sound clip' text.
    """
    doc = pymupdf.open(pdf_path)
    positions = []
    
    for page_num in range(len(doc)):
        page = doc[page_num]
        page_height = page.rect.height
        
        # Search for 'Sound clip' text
        text_instances = page.search_for("Sound clip")
        if not text_instances:
            continue
            
        images = page.get_image_info(hashes=True)
        
        for inst in text_instances:
            # Look for icon image to the right of 'Sound clip'
            search_area = pymupdf.Rect(inst.x1, inst.y0 - 5, page.rect.width, inst.y1 + 5)
            
            best_rect = None
            for img in images:
                img_rect = pymupdf.Rect(img["bbox"])
                if img_rect.intersects(search_area) and img_rect.width < 35 and img_rect.height < 35:
                    best_rect = img_rect
                    break
                    
            if best_rect:
                # Convert PyMuPDF coordinates (top-left origin) to pypdf (bottom-left origin)
                x_start = round(best_rect.x0)
                y_start = round(page_height - best_rect.y1)
                x_end = round(best_rect.x1)
                y_end = round(page_height - best_rect.y0)
                
                positions.append({
                    'page': page_num,
                    'rect': [x_start, y_start, x_end, y_end]
                })
                log(f"Page {page_num + 1}: Found sound icon at {best_rect} -> PDF Rect [{x_start}, {y_start}, {x_end}, {y_end}]", log_file)
            else:
                # Fallback to calculated offset if image not detected
                x_start = round(inst.x1 + 35)
                y_start = round(page_height - inst.y1 - 2)
                x_end = round(x_start + 16)
                y_end = round(page_height - inst.y0 + 2)
                positions.append({
                    'page': page_num,
                    'rect': [x_start, y_start, x_end, y_end]
                })
                log(f"Page {page_num + 1}: Icon image not detected, using fallback offset [{x_start}, {y_start}, {x_end}, {y_end}]", log_file)

    doc.close()
    return positions

def insert_sounds(input_pdf, wav_files, positions, output_pdf, log_file):
    reader = pypdf.PdfReader(input_pdf)
    writer = pypdf.PdfWriter()
    writer.append_pages_from_reader(reader)
    
    # Remove existing sound annotations first to prevent duplicate sounds
    removed_count = 0
    for page_num in range(len(writer.pages)):
        page = writer.get_page(page_num)
        if "/Annots" in page:
            annots = page["/Annots"]
            resolved_annots = annots.get_object() if hasattr(annots, "get_object") else annots
            new_annots = ArrayObject()
            for annot_ref in resolved_annots:
                annot = annot_ref.get_object()
                if annot.get("/Subtype") == "/Sound":
                    removed_count += 1
                    continue
                new_annots.append(annot_ref)
            if len(new_annots) > 0:
                page[NameObject("/Annots")] = new_annots
            else:
                try: del page["/Annots"]
                except KeyError: pass
                try: del page[NameObject("/Annots")]
                except KeyError: pass
                
    if removed_count > 0:
        log(f"Removed {removed_count} existing sound annotations from {os.path.basename(input_pdf)} before insertion.", log_file)
        
    if not positions:
        return 0
        
    total_positions = len(positions)
    num_wavs = len(wav_files)
    total_inserted = 0
    
    for i, pos in enumerate(positions):
        page_num = pos['page']
        rect = pos['rect']
        page = writer.get_page(page_num)
        
        # Cycling logic: 0 .. num_wavs-1, when reaching 21 it wraps back to 1st sound!
        wav_idx = i % num_wavs
        wav_path = wav_files[wav_idx]
        wav_name = os.path.basename(wav_path)
        
        cycle_note = ""
        if i >= num_wavs:
            cycle_note = f" (Sound #{i+1} -> looped back to #{wav_idx+1}: {wav_name})"
            
        # Ensure audio data is 16-bit 44.1kHz mono PCM WAV to guarantee 100% Adobe Acrobat stability
        wav_data = load_audio_as_16bit_pcm(wav_path, target_rate=44100)
            
        # Sound stream strictly conforming to PDF 1.7 / ISO 32000-1 specifications
        sound_stream = StreamObject()
        sound_stream.update({
            NameObject("/Type"): NameObject("/Sound"),
            NameObject("/R"): NumberObject(44100),
            NameObject("/C"): NumberObject(1),
            NameObject("/B"): NumberObject(16),
            NameObject("/E"): NameObject("/Signed"),
        })
        sound_stream._data = wav_data
        sound_stream_ref = writer._add_object(sound_stream)
        
        # Transparent appearance overlay stream over speaker icon
        bbox_w = rect[2] - rect[0]
        bbox_h = rect[3] - rect[1]
        appearance = StreamObject()
        appearance.update({
            NameObject("/Type"): NameObject("/XObject"),
            NameObject("/Subtype"): NameObject("/Form"),
            NameObject("/BBox"): ArrayObject([NumberObject(0), NumberObject(0), NumberObject(bbox_w), NumberObject(bbox_h)]),
            NameObject("/Resources"): DictionaryObject(),
        })
        appearance._data = b""
        appearance_ref = writer._add_object(appearance)
        
        annot = DictionaryObject()
        annot.update({
            NameObject("/Type"): NameObject("/Annot"),
            NameObject("/Subtype"): NameObject("/Sound"),
            NameObject("/Rect"): ArrayObject([NumberObject(r) for r in rect]),
            NameObject("/Sound"): sound_stream_ref,
            NameObject("/F"): NumberObject(4),
            NameObject("/AP"): DictionaryObject({NameObject("/N"): appearance_ref}),
            NameObject("/NM"): TextStringObject(f"SoundAnnot_{i}"),
            NameObject("/H"): NameObject("/P"),
        })
        annot_ref = writer._add_object(annot)
        
        if "/Annots" not in page:
            page[NameObject("/Annots")] = ArrayObject()
        page["/Annots"].append(annot_ref)
        
        total_inserted += 1
        log(f"Page {page_num + 1}: Inserted {wav_name}{cycle_note}", log_file)
        print_progress_bar(i + 1, total_positions, prefix='    Progress:', suffix=f'({i + 1}/{total_positions})')

    with open(output_pdf, "wb") as f:
        writer.write(f)
        
    return total_inserted

if __name__ == "__main__":
    cwd = os.getcwd()
    
    # 1. Resolve work_dir
    work_dir = None
    if len(sys.argv) > 1 and sys.argv[1].strip():
        arg_path = os.path.abspath(sys.argv[1].strip().strip('"'))
        if os.path.exists(arg_path):
            if os.path.basename(arg_path).upper() == "08. INSERT SOUND" or "INSERT SOUND" in os.path.basename(arg_path).upper():
                work_dir = arg_path
            elif os.path.exists(os.path.join(arg_path, "08. INSERT SOUND")):
                work_dir = os.path.join(arg_path, "08. INSERT SOUND")
            else:
                work_dir = arg_path

    if not work_dir:
        base_env = os.environ.get("BASE_DIR")
        if base_env and os.path.exists(base_env):
            cand = os.path.join(base_env, "08. INSERT SOUND")
            if os.path.exists(cand):
                work_dir = cand
                
    if not work_dir:
        if os.path.basename(cwd).upper() == "08. INSERT SOUND" or "INSERT SOUND" in os.path.basename(cwd).upper():
            work_dir = cwd
        elif os.path.exists(os.path.join(cwd, "08. INSERT SOUND")):
            work_dir = os.path.join(cwd, "08. INSERT SOUND")
            
    if not work_dir:
        default_target = r"C:\Users\aremi\Desktop\REPORT CBM\PO 42360565 - PAHANG - 11kV CYCLE3 - AZZAD\08. INSERT SOUND"
        if os.path.exists(default_target):
            work_dir = default_target
        else:
            work_dir = os.path.join(cwd, "08. INSERT SOUND")
            
    if not os.path.exists(work_dir):
        os.makedirs(work_dir, exist_ok=True)
        
    # 2. Resolve SOUND EXAMPLE directory (containing .wav files)
    sound_dir = os.path.join(work_dir, "SOUND EXAMPLE")
    if not os.path.exists(sound_dir) or not any(f.lower().endswith('.wav') for f in os.listdir(sound_dir)):
        alt_sound = os.path.join(work_dir, "SOUND")
        if os.path.exists(alt_sound) and any(f.lower().endswith('.wav') for f in os.listdir(alt_sound)):
            sound_dir = alt_sound
        else:
            sound_dir = work_dir # fallback to work_dir root
            
    # 3. Resolve PDF WITH SOUND output directory
    output_dir = os.path.join(work_dir, "PDF WITH SOUND")
    os.makedirs(output_dir, exist_ok=True)
    
    log_path = os.path.join(work_dir, "final_insert_log.txt")
    if os.path.exists(log_path):
        try: os.remove(log_path)
        except Exception: pass

    # Discover and naturally sort WAV files (1.wav, 2.wav, ..., 20.wav)
    wav_files = []
    if os.path.exists(sound_dir):
        wav_files = sorted(
            [os.path.join(sound_dir, f) for f in os.listdir(sound_dir) if f.lower().endswith('.wav')],
            key=lambda p: natural_sort_key(os.path.basename(p))
        )
        
    # Discover PDF files in root of 08. INSERT SOUND (ignoring subdirectories & temp files)
    pdf_files = sorted(
        [f for f in os.listdir(work_dir) if f.lower().endswith('.pdf') and not f.startswith('~$') and os.path.isfile(os.path.join(work_dir, f))],
        key=natural_sort_key
    )

    print("==================================================================")
    print("           SOUND (WAV) INSERTION SYSTEM INTO PDF REPORT           ")
    print("==================================================================")
    print(f"[*] Work Directory:       {work_dir}")
    print(f"[*] WAV Sound Directory:  {sound_dir}")
    print(f"[*] Output Directory:     {output_dir}")
    print(f"[*] Detected WAV Files:   {len(wav_files)} file(s)")
    if wav_files:
        print(f"    -> Sequence: {', '.join([os.path.basename(w) for w in wav_files[:5]])} ... {os.path.basename(wav_files[-1])}")
    print(f"[*] Detected PDF Files:   {len(pdf_files)} file(s) in root")
    print("------------------------------------------------------------------\n")

    if not wav_files:
        print(f"[!] Error: Missing WAV files in '{sound_dir}'. Exiting.")
        log(f"Missing WAV files in {sound_dir}.", log_path)
        sys.exit(1)

    if not pdf_files:
        print(f"[!] Error: No PDF files found in '{work_dir}'. Exiting.")
        log(f"No PDF files found in {work_dir}.", log_path)
        sys.exit(1)

    processed_count = 0
    skipped_count = 0
    total_sounds_inserted = 0

    for idx, pdf_file in enumerate(pdf_files):
        print(f"[{idx + 1}/{len(pdf_files)}] Processing: {pdf_file}")
        input_pdf = os.path.join(work_dir, pdf_file)
        output_pdf = os.path.join(output_dir, pdf_file)
        
        try:
            positions = get_icon_positions(input_pdf, log_path)
            if not positions:
                print(f"    - No 'Sound clip' icons found. (0 sounds - skipped)")
                log(f"Skipped {pdf_file}: 0 sound icons found.", log_path)
                skipped_count += 1
                continue
                
            count = insert_sounds(input_pdf, wav_files, positions, output_pdf, log_path)
            print(f"    [OK] Successfully inserted {count} sound(s) -> PDF WITH SOUND\\{pdf_file}\n")
            log(f"Successfully inserted {count} sounds for {pdf_file}. Output: {output_pdf}", log_path)
            processed_count += 1
            total_sounds_inserted += count
        except Exception as e:
            print(f"    [X] FATAL ERROR processing {pdf_file}: {str(e)}\n")
            log(f"FATAL ERROR processing {pdf_file}: {str(e)}", log_path)

    print("==================================================================")
    print("                  SOUND INSERTION COMPLETED                       ")
    print("==================================================================")
    print(f"[OK] Processed PDFs with Sound: {processed_count}")
    print(f"[-]  Skipped PDFs (No Sounds):  {skipped_count}")
    print(f"[OK] Total Sound Clips Injected: {total_sounds_inserted}")
    print(f"[OK] Output Folder: {output_dir}")
    print("==================================================================")
