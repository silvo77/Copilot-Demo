from dataclasses import dataclass
from typing import List, Optional, Tuple
import pandas as pd
from pathlib import Path
import argparse
import sys
import logging
import re
from datetime import datetime
import os

# Configura il logging
logs_dir = "logs"
if not os.path.exists(logs_dir):
    os.makedirs(logs_dir)

log_filename = os.path.join(logs_dir, f"lecture_parser_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_filename, encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)

# Verifica che il logging sia configurato correttamente
logging.info("Inizializzazione del sistema di logging")
logging.debug("Test debug message")

@dataclass
class Lecture:
    """Rappresenta una singola lezione del webinar"""
    type: str           # Tipo (Video, Doc, Section, etc.)
    title: str         # Titolo della lezione
    duration: int      # Durata in minuti
    lecture_number: int # Numero progressivo della lezione
    section_number: int # Numero della sezione di appartenenza
    start_time: float  # Inizio stimato in minuti dall'inizio del corso (può contenere frazioni per i secondi)
    end_time: float    # Fine stimata in minuti dall'inizio del corso (può contenere frazioni per i secondi)
    trovato: bool = False  # Indica se il titolo è stato trovato nel video

    @property
    def time_range(self) -> str:
        """Restituisce il range temporale in formato HH:MM:SS-HH:MM:SS"""
        def mins_to_hhmmss(mins: float) -> str:
            total_seconds = int(mins * 60)
            hours = total_seconds // 3600
            minutes = (total_seconds % 3600) // 60
            seconds = total_seconds % 60
            return f"{hours:02d}:{minutes:02d}:{seconds:02d}"
        
        return f"{mins_to_hhmmss(self.start_time)}-{mins_to_hhmmss(self.end_time)}"

class CourseParser:
    def __init__(self):
        self.lectures: List[Lecture] = []
        self._current_section = 0
        self._current_lecture = 0
    
    def parse_excel(self, file_path: str | Path) -> List[Lecture]:
        """Legge il file Excel e crea la lista delle lezioni"""
        logging.info(f"Inizio parsing del file Excel: {file_path}")
        def parse_duration(duration_str: str) -> int:
            """Converte una stringa di durata in minuti"""
            if pd.isna(duration_str):
                return 0
            
            duration_str = str(duration_str).strip().lower()
            if not duration_str:
                return 0

            # Se contiene '|', prendi la parte dopo il '|'
            if '|' in duration_str:
                duration_str = duration_str.split('|')[-1].strip()

            total_minutes = 0
            
            # Cerca ore
            hour_match = re.search(r'(\d+)\s*hr?', duration_str)
            if hour_match:
                total_minutes += int(hour_match.group(1)) * 60
            
            # Cerca minuti
            min_match = re.search(r'(\d+)\s*min', duration_str)
            if min_match:
                total_minutes += int(min_match.group(1))
            
            # Se non abbiamo trovato né ore né minuti, cerca solo numeri
            if total_minutes == 0:
                numbers = re.findall(r'\d+', duration_str)
                if numbers:
                    # Assume che il primo numero sia minuti
                    total_minutes = int(numbers[0])
            
            if total_minutes == 0:
                logging.warning(f"Impossibile interpretare la durata: '{duration_str}', usando 0 minuti")
            
            return total_minutes

        # Leggi solo le colonne necessarie: A, B e C
        df = pd.read_excel(
            file_path, 
            usecols=[0, 1, 2],  # A=0, B=1, C=2
            names=['type', 'title', 'duration']
        )
        
        # Processa ogni riga
        for _, row in df.iterrows():
            try:
                # Converti la durata usando la nuova funzione
                duration = parse_duration(row['duration'])
                logging.debug(f"Durata convertita: '{row['duration']}' → {duration} minuti")

                if row['type'].lower() == 'section':
                    self._current_section += 1
                    logging.debug(f"Nuova sezione trovata: {row['title']}")
                else:
                    self._current_lecture += 1
                    lecture = Lecture(
                        type=row['type'],
                        title=row['title'],
                        duration=duration,
                        lecture_number=self._current_lecture,
                        section_number=self._current_section,
                        start_time=0,  # Sarà calcolato dopo
                        end_time=0     # Sarà calcolato dopo
                    )
                    self.lectures.append(lecture)
                    logging.debug(f"Aggiunta lezione: {lecture.title} (durata: {duration}m)")
            except Exception as e:
                logging.error(f"Errore nel processare la riga: {row}")
                logging.error(f"Errore: {str(e)}")
                print(f"Errore nel processare la riga: {row}")
                print(f"Errore: {str(e)}")
        
        logging.info(f"Parsing completato: {len(self.lectures)} lezioni trovate in {self._current_section} sezioni")
        return self.lectures
    
    def calculate_times(self, start_idx: Optional[int] = None, end_idx: Optional[int] = None) -> None:
        """Calcola i tempi di inizio e fine per il range di lezioni specificato"""
        logging.info(f"Calcolo tempi per il range {start_idx or 0} - {end_idx or 'fine'}")
        sorted_lectures = sorted(self.lectures, key=lambda x: (x.section_number, x.lecture_number))
        
        if start_idx is None:
            start_idx = 0
        if end_idx is None:
            end_idx = len(sorted_lectures)
        
        # Assicurati che gli indici siano validi
        start_idx = max(0, min(start_idx, len(sorted_lectures)))
        end_idx = max(start_idx, min(end_idx, len(sorted_lectures)))
        
        # Calcola i tempi solo per il range specificato
        current_time = 0
        for i in range(start_idx, end_idx):
            lecture = sorted_lectures[i]
            lecture.start_time = current_time
            lecture.end_time = current_time + lecture.duration
            current_time = lecture.end_time
    
    def print_summary(self, start_idx: Optional[int] = None, end_idx: Optional[int] = None) -> None:
        """Stampa un riepilogo sintetico delle lezioni"""
        if not self.lectures:
            print("Nessuna lezione trovata")
            return
        
        sorted_lectures = sorted(self.lectures, key=lambda x: (x.section_number, x.lecture_number))
        
        if start_idx is None:
            start_idx = 0
        if end_idx is None:
            end_idx = len(sorted_lectures)
        
        # Assicurati che gli indici siano validi
        start_idx = max(0, min(start_idx, len(sorted_lectures)))
        end_idx = max(start_idx, min(end_idx, len(sorted_lectures)))
        
        # Stampa intestazione
        print("\nS.## L.## Start       End         Title                                          Duration")
        print("-" * 95)
        
        total_duration = 0
        for lecture in sorted_lectures[start_idx:end_idx]:
            total_duration += lecture.duration
            # Tronca il titolo se troppo lungo
            title = lecture.title[:40] + "..." if len(lecture.title) > 40 else lecture.title.ljust(43)
            # Converti i tempi in formato HH:MM:SS
            start_secs = int(lecture.start_time * 60)
            end_secs = int(lecture.end_time * 60)
            start = f"{start_secs // 3600:02d}:{(start_secs % 3600) // 60:02d}:{start_secs % 60:02d}"
            end = f"{end_secs // 3600:02d}:{(end_secs % 3600) // 60:02d}:{end_secs % 60:02d}"
            # Arrotonda la durata all'intero più vicino per la visualizzazione
            duration_int = int(round(lecture.duration))
            print(f"{lecture.section_number:2d}.{lecture.lecture_number:2d}  {start}    {end}    {title} {duration_int:3d}m")
        
        print("-" * 75)
        hours = int(total_duration) // 60
        minutes = int(total_duration) % 60
        n_lectures = end_idx - start_idx
        print(f"Total: {n_lectures} lectures, {hours}h {minutes}m total duration")
    
    def set_start_time(self, start_time_str: str, start_idx: Optional[int] = None, end_idx: Optional[int] = None) -> None:
        """Imposta l'orario di inizio per il range specificato
        
        Args:
            start_time_str: Orario di inizio nel formato "HH:MM[:SS]"
            start_idx: Indice della prima lezione (1-based, opzionale)
            end_idx: Indice dell'ultima lezione (1-based, opzionale)
        """
        logging.info(f"Impostazione orario di inizio {start_time_str} per il range {start_idx or 1} - {end_idx or 'fine'}")
        # Converti HH:MM[:SS] in minuti
        time_parts = start_time_str.split(':')
        if len(time_parts) == 3:
            hours, minutes, seconds = map(int, time_parts)
        elif len(time_parts) == 2:
            hours, minutes = map(int, time_parts)
            seconds = 0
        else:
            raise ValueError("Il formato dell'orario deve essere HH:MM o HH:MM:SS")
        
        start_time = hours * 60 + minutes + seconds / 60
        
        sorted_lectures = sorted(self.lectures, key=lambda x: (x.section_number, x.lecture_number))
        
        if start_idx is None:
            start_idx = 0
        else:
            start_idx = max(0, min(start_idx - 1, len(sorted_lectures)))
            
        if end_idx is None:
            end_idx = len(sorted_lectures)
        else:
            end_idx = max(start_idx, min(end_idx, len(sorted_lectures)))
        
        # Calcola i nuovi orari partendo dall'orario specificato
        current_time = start_time
        for i in range(start_idx, end_idx):
            lecture = sorted_lectures[i]
            old_start = lecture.start_time
            old_end = lecture.end_time
            lecture.start_time = current_time
            lecture.end_time = current_time + lecture.duration
            current_time = lecture.end_time
            logging.debug(f"Aggiornata lezione {lecture.section_number}.{lecture.lecture_number} '{lecture.title}': "
                        f"{old_start}-{old_end} → {lecture.start_time}-{lecture.end_time}")
    
    def set_end_time(self, end_time_str: str, start_idx: Optional[int] = None, end_idx: Optional[int] = None) -> None:
        """Imposta l'orario di fine per una lezione specifica e aggiorna gli orari"""
        logging.info(f"Impostazione orario di fine {end_time_str} per la lezione {start_idx or 'ultima'}")
        # Converti HH:MM[:SS] in minuti
        time_parts = end_time_str.split(':')
        if len(time_parts) == 3:
            hours, minutes, seconds = map(int, time_parts)
        elif len(time_parts) == 2:
            hours, minutes = map(int, time_parts)
            seconds = 0
        else:
            raise ValueError("Il formato dell'orario deve essere HH:MM o HH:MM:SS")
            
        end_time = hours * 60 + minutes + seconds / 60
        
        sorted_lectures = sorted(self.lectures, key=lambda x: (x.section_number, x.lecture_number))
        
        if start_idx is None:
            start_idx = len(sorted_lectures) - 1  # Ultima lezione se non specificato
        else:
            start_idx = max(0, min(start_idx - 1, len(sorted_lectures) - 1))
        
        # Imposta il nuovo orario di fine per la lezione specificata e aggiorna la sua durata
        target_lecture = sorted_lectures[start_idx]
        old_end = target_lecture.end_time
        old_duration = target_lecture.duration
        target_lecture.end_time = end_time
        target_lecture.duration = end_time - target_lecture.start_time
        logging.debug(f"Aggiornata lezione {target_lecture.section_number}.{target_lecture.lecture_number} "
                    f"'{target_lecture.title}': fine {old_end} → {end_time}, durata {old_duration} → {target_lecture.duration}")
        
        # Aggiorna gli orari delle lezioni successive
        current_time = end_time
        for i in range(start_idx + 1, len(sorted_lectures)):
            lecture = sorted_lectures[i]
            old_start = lecture.start_time
            old_end = lecture.end_time
            lecture.start_time = current_time
            lecture.end_time = current_time + lecture.duration
            current_time = lecture.end_time
            logging.debug(f"Aggiornata lezione successiva {lecture.section_number}.{lecture.lecture_number} "
                        f"'{lecture.title}': {old_start}-{old_end} → {lecture.start_time}-{lecture.end_time}")

def parse_range(range_str: str) -> Tuple[Optional[int], Optional[int]]:
    """Converte una stringa range (es. '1-5' o '3-' o '-7') in una tupla di indici"""
    if not range_str:
        return None, None
    
    # Gestisce il caso in cui il range non sia nel formato corretto
    if '-' not in range_str:
        range_str = f"{range_str}-"
    
    parts = range_str.split('-')
    if len(parts) != 2:
        raise ValueError("Il range deve essere nel formato 'start-end' (es. '1-5' o '3-' o '-7')")
    
    start = int(parts[0]) if parts[0] else None  # Converti in 1-based
    end = int(parts[1]) if parts[1] else start  # end è già inclusivo
    
    return start, end

def main():
    parser = argparse.ArgumentParser(description="Parser per contenuti del corso da file Excel")
    parser.add_argument('excel_file', help='Il file Excel da processare')
    parser.add_argument('--range', '-r', help='Range di lezioni da mostrare (es. "1-5" o "3-" o "-7")')
    parser.add_argument('--verbose', '-v', action='store_true', help='Abilita output dettagliato')
    
    try:
        args = parser.parse_args()
        
        # Imposta il livello di logging in base all'argomento verbose
        if args.verbose:
            logging.getLogger().setLevel(logging.DEBUG)
        
        logging.info("Avvio elaborazione")
        course_parser = CourseParser()
        course_parser.parse_excel(args.excel_file)
        
        start_idx, end_idx = parse_range(args.range) if args.range else (None, None)
        if args.range:
            logging.info(f"Range specificato: {args.range} → indici {start_idx}-{end_idx}")
        
        # Prima calcola i tempi per il range specificato
        course_parser.calculate_times(start_idx, end_idx)
        # Poi stampa il sommario
        course_parser.print_summary(start_idx, end_idx)
        
    except Exception as e:
        logging.error(f"Errore durante l'esecuzione: {e}", exc_info=True)
        print(f"Errore: {e}", file=sys.stderr)
        sys.exit(1)
    
    logging.info("Elaborazione completata")
