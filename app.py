import argparse
import csv
import sys
from pathlib import Path
from typing import List, Tuple, Optional
import logging
from datetime import datetime

from lecture_parser import CourseParser, Lecture
from find_text_in_video import find_text_in_video, setup_logging, seconds_to_hms
from chapter_manager import ChapterManager

def add_video_chapters(video_file: str, lectures: List[Lecture], timestamps: List[Tuple[float, float]]):
    """
    Aggiunge i capitoli al file video usando i timestamp trovati.
    
    Args:
        video_file: Percorso del file video
        lectures: Lista delle lezioni
        timestamps: Lista di tuple (inizio, fine) per ogni lezione
    """
    chapter_manager = ChapterManager()
    # Converti le lezioni in capitoli
    for lecture, (start, end) in zip(lectures, timestamps):
        if lecture.trovato:  # Aggiungi solo le lezioni trovate come capitoli
            setattr(lecture, 'chapter', True)
            setattr(lecture, 'chapter_start_time', start)
            setattr(lecture, 'chapter_end_time', int(end * 1000))  # Converti in millisecondi
    
    # Imposta le sessioni nel chapter manager
    chapter_manager.sessions = lectures
    
    # Esporta i metadati e aggiungi i capitoli
    metadata_file = "FFMETADATAFILE.txt"
    chapter_manager.export_chapters_metadata(video_file, metadata_file)
    chapter_manager.add_chapters_to_video_file(video_file, metadata_file)
    
    print(f"\nCapitoli aggiunti al video: {video_file}")

def find_lecture_timestamps(video_file: str, lectures: List[Lecture], search_window: float = 300, 
                        crop_area: Optional[Tuple[int, int, int, int]] = None,
                        truncate_length: Optional[int] = None,
                        save_frames: bool = False,
                        strip_prefix: bool = False) -> List[Tuple[float, float]]:
    """
    Cerca i timestamp di inizio e fine di ogni lezione nel video.
    
    Args:
        video_file: Percorso del file video
        lectures: Lista delle lezioni da cercare
        search_window: Finestra di ricerca in secondi intorno al timestamp stimato
        left: Margine sinistro del ritaglio
        top: Margine superiore del ritaglio
        right: Margine destro del ritaglio
        bottom: Margine inferiore del ritaglio
    
    Returns:
        Lista di tuple (timestamp_inizio, timestamp_fine) per ogni lezione
    """
    timestamps = []
    last_end_time = 0  # Teniamo traccia dell'ultimo tempo di fine trovato
    last_lecture_idx = -1  # Indice dell'ultima lezione processata
    
    for i, lecture in enumerate(lectures):
        logging.info(f"Cercando lezione {i+1}/{len(lectures)}: {lecture.title}")
        
        # Per la prima lezione, inizia da zero
        # Per le altre, inizia dall'ultimo tempo di fine trovato
        search_start = last_end_time
        
        # Prepara il testo da cercare
        search_text = lecture.title
        # Rimuovi la parte iniziale fino al primo spazio se richiesto o se è una lezione di tipo doc
        if strip_prefix or lecture.type.lower() == "doc":
            # Rimuovi la parte iniziale fino al primo spazio (es: "150. " da "150. Introduction")
            search_text = lecture.title.split(" ", 1)[1] if " " in lecture.title else lecture.title
        
        # Tronca il testo se richiesto
        original_text = search_text
        if truncate_length is not None and len(search_text) > truncate_length:
            search_text = search_text[:truncate_length]
            logging.info(f"Testo troncato da '{original_text}' a '{search_text}'")
        search_text = search_text.strip()
        
        # Determina la finestra di ricerca
        current_window = search_window
        # Se non è la prima lezione e la lezione precedente è di tipo doc, estendi la finestra
        if i > 0 and lectures[i-1].type.lower() == "doc":
            current_window += 60  # Aggiungi 60 secondi alla finestra
            logging.info(f"Finestra di ricerca estesa di 1 minuto (lezione precedente di tipo doc)")
        
        # Cerca il titolo della lezione
        print(f"\nCercando inizio lezione: {lecture.title}")
        print(f"Testo ricercato: {search_text}")
        print(f"Finestra di ricerca: {seconds_to_hms(current_window)}")
        
        # Usa l'area di ritaglio se specificata
        if crop_area is not None:
            # Valida le percentuali
            if not all(0 <= x <= 100 for x in crop_area):
                print("Error: I valori di ritaglio devono essere percentuali tra 0 e 100")
                sys.exit(1)
            left, top, right, bottom = crop_area
            # Assicurati che left < right e top < bottom
            if left >= right or top >= bottom:
                print("Error: I valori LEFT/RIGHT e TOP/BOTTOM non sono validi")
                sys.exit(1)
            logging.info(f"Area di ritaglio specificata (percentuale): L={left}% T={top}% R={right}% B={bottom}%")
        else:
            logging.info("Nessuna area di ritaglio specificata, analizzando l'intero frame")
        
        # Passa crop_area e save_frames alla funzione find_text_in_video
        start_time, _, _ = find_text_in_video(
            video_file,
            search_start,
            current_window,
            search_text,
            crop_area=crop_area,
            save_frames=save_frames
        )
        
        # Aggiorna lo stato di ricerca della lezione
        if start_time is None:
            logging.warning(f"Inizio lezione non trovato per: {lecture.title}. Usando tempo stimato.")
            start_time = last_end_time  # Partiamo dalla fine della lezione precedente
            lecture.trovato = False
        else:
            lecture.trovato = True
            logging.info(f"Lezione trovata: {lecture.title}")
        
        # Se non è la prima lezione, usa questo tempo come fine della lezione precedente
        if last_lecture_idx >= 0:
            timestamps[last_lecture_idx] = (timestamps[last_lecture_idx][0], start_time)
            print(f"Aggiornato tempo fine lezione {last_lecture_idx + 1}: {seconds_to_hms(start_time)}")
        
        # Per l'ultima lezione, il tempo di fine sarà la durata stimata
        if i == len(lectures) - 1:
            end_time = start_time + (lecture.duration * 60)
        else:
            # Per le altre lezioni, il tempo di fine sarà temporaneo finché non troviamo l'inizio della prossima
            end_time = start_time + (lecture.duration * 60)  # Valore temporaneo
        
        timestamps.append((start_time, end_time))
        last_end_time = end_time  # Aggiorniamo l'ultimo tempo di fine
        last_lecture_idx = i
        logging.info(f"Prossima ricerca partirà da: {seconds_to_hms(last_end_time)}")
        print(f"Lezione {i+1}: {lecture.title}")
        print(f"  Inizio: {seconds_to_hms(start_time)}")
        print(f"  Fine: {seconds_to_hms(end_time)}")
        print(f"  Durata: {seconds_to_hms(end_time-start_time)}")
    
    return timestamps

def export_to_csv(lectures: List[Lecture], timestamps: List[Tuple[float, float]], output_file: str):
    """
    Esporta le lezioni e i loro timestamp in un file CSV.
    """
    with open(output_file, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(['Sezione', 'Lezione', 'Titolo', 'Inizio', 'Fine', 'Durata', 'Trovato'])
        
        for lecture, (start, end) in zip(lectures, timestamps):
            writer.writerow([
                lecture.section_number,
                lecture.lecture_number,
                lecture.title,
                seconds_to_hms(start),
                seconds_to_hms(end),
                seconds_to_hms(end-start),
                'Yes' if lecture.trovato else 'No'
            ])
    
    print(f"\nDati esportati in: {output_file}")

def main():
    parser = argparse.ArgumentParser(description="Trova i timestamp delle lezioni in un video")
    parser.add_argument('video', help='File video da analizzare')
    parser.add_argument('excel', help='File Excel con la struttura del corso')
    group = parser.add_mutually_exclusive_group()
    group.add_argument('--range', '-r', help='Range di lezioni da processare (es. "1-5" o "3-" o "-7")')
    group.add_argument('--section', '-s', help='Range di sezioni da processare (es. "1-2" o "2-" o "-3")')
    parser.add_argument('--output', '-o', help='File CSV di output (default: timestamps.csv)')
    parser.add_argument('--window', '-w', type=float, default=90,
                       help='Finestra di ricerca in secondi intorno al timestamp stimato (default: 90)')
    parser.add_argument('--verbose', '-v', action='store_true', help='Abilita output dettagliato')
    
    # Aggiungi parametro per l'area di ritaglio (in percentuale)
    parser.add_argument('--crop', type=float, nargs=4, metavar=('LEFT', 'TOP', 'RIGHT', 'BOTTOM'),
                       help='Area di ritaglio del video in percentuale (es. --crop 10 5 90 15)')
    
    # Aggiungi parametro per troncare il testo di ricerca
    parser.add_argument('--truncate', '-t', type=int, metavar='N',
                       help='Tronca il testo di ricerca ai primi N caratteri')
    parser.add_argument('--save-frames', action='store_true',
                       help='Salva le immagini dei frame dove viene trovato il testo')
    parser.add_argument('--strip-prefix', action='store_true',
                       help='Rimuovi il testo fino al primo spazio per tutte le lezioni')
    
    args = parser.parse_args()
    
    # Setup logging
    setup_logging(args.video)
    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)
    
    # Parse Excel
    course_parser = CourseParser()
    lectures = course_parser.parse_excel(args.excel)
    
    # Ordina le lezioni per sezione e numero
    lectures = sorted(lectures, key=lambda x: (x.section_number, x.lecture_number))

    # Applica il range di lezioni o sezioni se specificato
    if args.range or args.section:
        from lecture_parser import parse_range
        if args.range:
            # Filtra per range di lezioni
            start_idx, end_idx = parse_range(args.range)
            lectures = lectures[start_idx+1:end_idx+1]
        else:
            # Filtra per range di sezioni
            start_section, end_section = parse_range(args.section)
            print(f"Filtrando lezioni tra le sezioni {start_section} e {end_section}")
            lectures = [l for l in lectures if 
                       l.section_number is not None and 
                       start_section <= l.section_number <= end_section]
            lecture_numbers = [lecture.lecture_number for lecture in lectures]
            print(f"Lezioni trovate: {lecture_numbers}")  
            if not lectures:
                print(f"Nessuna lezione trovata nelle sezioni {args.section}")
                sys.exit(1)
    
    # Calcola i tempi stimati
    course_parser.calculate_times()
    
    # Trova i timestamp effettivi nel video
    timestamps = find_lecture_timestamps(
        args.video, 
        lectures, 
        args.window,
        crop_area=args.crop,
        truncate_length=args.truncate,
        save_frames=args.save_frames,
        strip_prefix=args.strip_prefix
    )
    
    # Esporta in CSV
    output_file = args.output or 'timestamps.csv'
    export_to_csv(lectures, timestamps, output_file)
    
    # Prepara il sommario delle lezioni trovate e non trovate
    found_lectures = [(i, l) for i, l in enumerate(lectures) if l.trovato]
    not_found_lectures = [(i, l) for i, l in enumerate(lectures) if not l.trovato]
    
    print("\n=== Sommario della ricerca ===")
    print(f"Totale lezioni processate: {len(lectures)}")
    print(f"Lezioni trovate: {len(found_lectures)}")
    if found_lectures:
        found_numbers = [l.lecture_number for _, l in found_lectures]
        print(f"  Numeri lezioni trovate: {', '.join(map(str, found_numbers))}")
    print(f"Lezioni non trovate: {len(not_found_lectures)}")
    if not_found_lectures:
        not_found_numbers = [l.lecture_number for _, l in not_found_lectures]
        print(f"  Numeri lezioni non trovate: {', '.join(map(str, not_found_numbers))}")
    print("=" * 30)
    if not_found_lectures:
        print("\nLezioni non trovate:")
        for i, lecture in not_found_lectures:
            start_time = timestamps[i][0]  # Prendi il timestamp di inizio usato per la ricerca
            print(f"{i+1}. {lecture.title}")
            print(f"   La ricerca è stata effettuata a partire da: {seconds_to_hms(start_time)}")
            while True:
                try:
                    response = input(f"Inserisci il timestamp manualmente (formato HH:MM:SS) o premi Invio per saltare: ").strip()
                    if not response:
                        break
                    
                    # Converti il timestamp in secondi
                    parts = response.split(':')
                    if len(parts) != 3:
                        print("Formato non valido. Usa HH:MM:SS")
                        continue
                    
                    hours, minutes, seconds = map(float, parts)
                    timestamp = hours * 3600 + minutes * 60 + seconds
                    
                    # Aggiorna il timestamp e marca la lezione come trovata
                    timestamps[i] = (timestamp, timestamp + (lecture.duration * 60))
                    lecture.trovato = True
                    print(f"Timestamp impostato a {seconds_to_hms(timestamp)}")
                    break
                except ValueError:
                    print("Formato non valido. Usa HH:MM:SS")
        
        # Chiedi conferma prima di procedere
        response = input("\nProcedere con la creazione dei capitoli? (s/n): ").strip().lower()
        if response != 's':
            print("Operazione annullata")
            sys.exit(0)
    
    # Aggiungi i capitoli al video
    print("\nAggiunta capitoli al video...")
    add_video_chapters(args.video, lectures, timestamps)

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nOperazione interrotta dall'utente")
        sys.exit(1)
    except Exception as e:
        logging.error(f"Errore: {str(e)}", exc_info=True)
        print(f"Errore: {str(e)}", file=sys.stderr)
        sys.exit(1)
