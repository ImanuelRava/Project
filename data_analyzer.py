import pandas as pd
import collections
import os

class DataAnalyzer:
    def __init__(self):
        self.matches = [] 
        self.hero_win_rates = collections.defaultdict(lambda: {'wins': 0, 'total': 0})
        self.is_loaded = False
        self.load_local_data()

    def find_file(self, target_filename):
        if os.path.exists(target_filename): return target_filename
        cwd = os.getcwd()
        files = os.listdir(cwd)
        target_lower = target_filename.lower()
        for f in files:
            if f.lower() == target_lower: return f
        return None

    def load_local_data(self):
        file_path = 'Analyst.xlsx'
        actual_path = self.find_file(file_path)
        
        if actual_path:
            print(f"Found {actual_path}. Processing tournament data...")
            success, msg = self.process_file(actual_path)
            if success:
                print("Tournament data loaded successfully.")
                self.is_loaded = True
            else:
                print(f"Failed to load tournament data: {msg}")
        else:
            print("Analyst.xlsx not found.")

    def process_file(self, file_path):
        try:
            df_raw = pd.read_excel(file_path, sheet_name=0, header=None)
            
            # Find Header Row
            header_row_idx = None
            for i, row in df_raw.iterrows():
                if 'Ban 1' in row.values:
                    header_row_idx = i
                    break
            
            if header_row_idx is None: return False, "Header row not found."

            # Find Data Start Row
            start_row_idx = None
            for i in range(header_row_idx + 1, len(df_raw)):
                val = str(df_raw.iloc[i, 0]).strip() 
                if val and val != '*' and val != 'nan' and 'Tournament' not in val:
                    start_row_idx = i
                    break
            
            if start_row_idx is None: return False, "Data rows not found."

            df = df_raw.iloc[start_row_idx:].reset_index(drop=True)
            
            # Map Columns
            header_row = df_raw.iloc[header_row_idx]
            header_row_norm = header_row.astype(str).str.strip().str.lower()
            ban1_cols = header_row_norm[header_row_norm == 'ban 1'].index.tolist()
            
            if len(ban1_cols) < 2: return False, "Column mapping failed."

            col_blue_start = ban1_cols[0]
            col_red_start = ban1_cols[1]
            
            col_blue_team = col_blue_start - 1
            col_red_team = col_red_start - 1
            
            def clean_name(name):
                if pd.isna(name): return ""
                return str(name).strip()

            for index, row in df.iterrows():
                try:
                    tournament = str(row.iloc[0]).strip()
                    selected_map = str(row.iloc[1]).strip()
                    
                    blue_team = clean_name(row.iloc[col_blue_team])
                    red_team = clean_name(row.iloc[col_red_team])
                    
                    # ==========================================
                    # CORRECTED PHASE LOGIC
                    # We slice first (preserving position), then clean.
                    # ==========================================
                    
                    # Blue Bans (Indices 0-4)
                    raw_blue_bans = row.iloc[col_blue_start : col_blue_start + 5].tolist()
                    blue_bans_p1 = [clean_name(b) for b in raw_blue_bans[:3]]     # Slots 1, 2, 3
                    blue_bans_p2 = [clean_name(b) for b in raw_blue_bans[3:5]]   # Slots 4, 5
                    
                    # Red Bans (Indices 0-4)
                    raw_red_bans = row.iloc[col_red_start : col_red_start + 5].tolist()
                    red_bans_p1 = [clean_name(b) for b in raw_red_bans[:3]]
                    red_bans_p2 = [clean_name(b) for b in raw_red_bans[3:5]]
                    
                    # Blue Picks (Indices 0-4)
                    raw_blue_picks = row.iloc[col_blue_start + 5 : col_blue_start + 10].tolist()
                    blue_picks_p1 = [clean_name(p) for p in raw_blue_picks[:3]]
                    blue_picks_p2 = [clean_name(p) for p in raw_blue_picks[3:5]]

                    # Red Picks (Indices 0-4)
                    raw_red_picks = row.iloc[col_red_start + 5 : col_red_start + 10].tolist()
                    red_picks_p1 = [clean_name(p) for p in raw_red_picks[:3]]
                    red_picks_p2 = [clean_name(p) for p in raw_red_picks[3:5]]

                    blue_result = str(row.iloc[col_blue_start + 10]).strip().upper()
                    red_result = str(row.iloc[col_red_start + 10]).strip().upper()

                    # Store Match Data
                    self.matches.append({
                        'tournament': tournament,
                        'map': selected_map,
                        'blue_team': blue_team,
                        'red_team': red_team,
                        'blue_bans': {'p1': blue_bans_p1, 'p2': blue_bans_p2},
                        'blue_picks': {'p1': blue_picks_p1, 'p2': blue_picks_p2},
                        'blue_result': blue_result,
                        'red_bans': {'p1': red_bans_p1, 'p2': red_bans_p2},
                        'red_picks': {'p1': red_picks_p1, 'p2': red_picks_p2},
                        'red_result': red_result
                    })
                    
                    # Win rates use total picks (filter empty strings)
                    all_blue = [p for p in blue_picks_p1 + blue_picks_p2 if p]
                    all_red = [p for p in red_picks_p1 + red_picks_p2 if p]
                    
                    self._update_hero_stats(all_blue, blue_result)
                    self._update_hero_stats(all_red, red_result)

                except Exception:
                    continue

            return True, "Success"
            
        except Exception as e:
            return False, str(e)

    def _update_hero_stats(self, picks, result):
        for hero in picks:
            self.hero_win_rates[hero]['total'] += 1
            if result == 'WIN':
                self.hero_win_rates[hero]['wins'] += 1

    def get_unique_values(self, column, tournament_filter=None):
        if not self.matches: return []
        
        if column == 'team':
            teams = set()
            for m in self.matches:
                if tournament_filter and tournament_filter != "All" and m['tournament'] != tournament_filter:
                    continue
                teams.add(m['blue_team'])
                teams.add(m['red_team'])
            return sorted(list(teams))
            
        return sorted(list(set(m[column] for m in self.matches)))

    def get_hero_summary(self, side, tournament_filter=None, map_filter=None, team_filter=None):
        stats = collections.defaultdict(lambda: {
            'picks_p1': 0, 'picks_p2': 0, 
            'bans_p1': 0, 'bans_p2': 0, 
            'wins': 0, 'total_picks': 0
        })
        
        for match in self.matches:
            if tournament_filter and tournament_filter != "All" and match['tournament'] != tournament_filter: continue
            if map_filter and map_filter != "All" and match['map'] != map_filter: continue
            
            target_picks, target_bans, target_result = None, None, None
            
            if side == 'Blue':
                if team_filter and team_filter != "All" and match['blue_team'] != team_filter: continue
                target_picks = match['blue_picks']
                target_bans = match['blue_bans']
                target_result = match['blue_result']
            else: 
                if team_filter and team_filter != "All" and match['red_team'] != team_filter: continue
                target_picks = match['red_picks']
                target_bans = match['red_bans']
                target_result = match['red_result']

            # Aggregate Picks
            # Filter empty strings inside the loop to avoid counting blanks
            for hero in target_picks['p1']:
                if hero:
                    stats[hero]['picks_p1'] += 1
                    stats[hero]['total_picks'] += 1
                    if target_result == 'WIN': stats[hero]['wins'] += 1
            
            for hero in target_picks['p2']:
                if hero:
                    stats[hero]['picks_p2'] += 1
                    stats[hero]['total_picks'] += 1
                    if target_result == 'WIN': stats[hero]['wins'] += 1

            # Aggregate Bans
            for hero in target_bans['p1']:
                if hero: stats[hero]['bans_p1'] += 1
            
            for hero in target_bans['p2']:
                if hero: stats[hero]['bans_p2'] += 1
        
        # Build DataFrame
        data = []
        for hero, s in stats.items():
            wr = (s['wins'] / s['total_picks'] * 100) if s['total_picks'] > 0 else 0.0
            data.append({
                'Hero': hero,
                'Pick P1': s['picks_p1'],
                'Pick P2': s['picks_p2'],
                'Ban P1': s['bans_p1'],
                'Ban P2': s['bans_p2'],
                'Total Picks': s['total_picks'], 
                'Win Rate': wr
            })
        
        return pd.DataFrame(data)

    def get_hero_win_rate(self, hero_name):
        data = self.hero_win_rates.get(hero_name)
        if not data or data['total'] == 0: return None
        return (data['wins'] / data['total']) * 100