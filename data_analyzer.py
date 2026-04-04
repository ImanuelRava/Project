import pandas as pd
import collections
import os
import itertools

class DataAnalyzer:
    def __init__(self):
        self.matches = [] 
        self.hero_win_rates = collections.defaultdict(lambda: {'wins': 0, 'total': 0})
        # Data structures for Synergy and Counters
        self.synergy_map = collections.defaultdict(lambda: {'wins': 0, 'total': 0})
        # Counter Map now stores H2H stats: Key=(HeroA, HeroB), Val={'wins': 0, 'total': 0}
        # 'total' represents the total number of times Hero A played against Hero B.
        self.counter_map = collections.defaultdict(lambda: {'wins': 0, 'total': 0})
        
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
        script_dir = os.path.dirname(os.path.abspath(__file__))
        file_path = os.path.join(script_dir, 'Analyst.xlsx')
        
        if os.path.exists(file_path):
            print(f"Found {file_path}. Processing tournament data...")
            success, msg = self.process_file(file_path)
            if success:
                print("Tournament data loaded successfully.")
                self.is_loaded = True
                self._calculate_relationships()
                print("Hero relationships (Synergy/Counters) calculated.")
            else:
                print(f"Failed to load tournament data: {msg}")
        else:
            print(f"CRITICAL: File not found at {file_path}")

    def process_file(self, file_path):
        try:
            df_raw = pd.read_excel(file_path, sheet_name='MLBB Statistics', header=None)
            
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
                    
                    # Phase Logic
                    raw_blue_bans = row.iloc[col_blue_start : col_blue_start + 5].tolist()
                    blue_bans_p1 = [clean_name(b) for b in raw_blue_bans[:3]]
                    blue_bans_p2 = [clean_name(b) for b in raw_blue_bans[3:5]]
                    
                    raw_red_bans = row.iloc[col_red_start : col_red_start + 5].tolist()
                    red_bans_p1 = [clean_name(b) for b in raw_red_bans[:3]]
                    red_bans_p2 = [clean_name(b) for b in raw_red_bans[3:5]]
                    
                    raw_blue_picks = row.iloc[col_blue_start + 5 : col_blue_start + 10].tolist()
                    blue_picks_p1 = [clean_name(p) for p in raw_blue_picks[:3]]
                    blue_picks_p2 = [clean_name(p) for p in raw_blue_picks[3:5]]

                    raw_red_picks = row.iloc[col_red_start + 5 : col_red_start + 10].tolist()
                    red_picks_p1 = [clean_name(p) for p in raw_red_picks[:3]]
                    red_picks_p2 = [clean_name(p) for p in raw_red_picks[3:5]]

                    blue_result = str(row.iloc[col_blue_start + 10]).strip().upper()
                    red_result = str(row.iloc[col_red_start + 10]).strip().upper()

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

    def _calculate_relationships(self):
        print("Calculating Hero Relationships...")
        for match in self.matches:
            blue_picks = [h for h in match['blue_picks']['p1'] + match['blue_picks']['p2'] if h]
            red_picks = [h for h in match['red_picks']['p1'] + match['red_picks']['p2'] if h]
            blue_win = match['blue_result'] == 'WIN'
            red_win = match['red_result'] == 'WIN'
            
            winning_team = blue_picks if blue_win else red_picks
            losing_team = red_picks if blue_win else blue_picks
            
            # Synergy Logic (Unchanged - teammates share fate)
            for pair in itertools.combinations(winning_team, 2):
                key = tuple(sorted(pair))
                self.synergy_map[key]['total'] += 1
                self.synergy_map[key]['wins'] += 1
            for pair in itertools.combinations(losing_team, 2):
                key = tuple(sorted(pair))
                self.synergy_map[key]['total'] += 1
                # wins not incremented for losing pair
            
            # Counter Logic (Fixed - H2H Head to Head)
            # We iterate every winner against every loser
            for winner_hero in winning_team:
                for loser_hero in losing_team:
                    # 1. Update Winner's record vs Loser
                    # Key (Winner, Loser): "I beat this guy"
                    key_win = (winner_hero, loser_hero)
                    self.counter_map[key_win]['wins'] += 1
                    self.counter_map[key_win]['total'] += 1
                    
                    # 2. Update Loser's record vs Winner
                    # Key (Loser, Winner): "I lost to this guy"
                    key_loss = (loser_hero, winner_hero)
                    self.counter_map[key_loss]['total'] += 1
                    # wins not incremented

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
        stats = collections.defaultdict(lambda: {'picks_p1': 0, 'picks_p2': 0, 'bans_p1': 0, 'bans_p2': 0, 'wins': 0, 'total_picks': 0})
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
            for hero in target_bans['p1']:
                if hero: stats[hero]['bans_p1'] += 1
            for hero in target_bans['p2']:
                if hero: stats[hero]['bans_p2'] += 1
        
        data = []
        for hero, s in stats.items():
            wr = (s['wins'] / s['total_picks'] * 100) if s['total_picks'] > 0 else 0.0
            data.append({'Hero': hero, 'Pick P1': s['picks_p1'], 'Pick P2': s['picks_p2'], 'Ban P1': s['bans_p1'], 'Ban P2': s['bans_p2'], 'Total Picks': s['total_picks'], 'Win Rate': wr})
        return pd.DataFrame(data)

    def get_hero_win_rate(self, hero_name):
        data = self.hero_win_rates.get(hero_name)
        if not data or data['total'] == 0: return None
        return (data['wins'] / data['total']) * 100

    def get_synergy_for_hero(self, hero_name):
        partners = collections.defaultdict(lambda: {'wins': 0, 'total': 0})
        for (h1, h2), stats in self.synergy_map.items():
            if h1 == hero_name: partners[h2]['wins'] += stats['wins']; partners[h2]['total'] += stats['total']
            elif h2 == hero_name: partners[h1]['wins'] += stats['wins']; partners[h1]['total'] += stats['total']
        
        result = []
        for partner, stats in partners.items():
            if stats['total'] >= 2:
                wr = (stats['wins'] / stats['total']) * 100
                result.append({'hero': partner, 'win_rate': wr, 'matches': stats['total']})
        result.sort(key=lambda x: (x['win_rate'], x['matches']), reverse=True)
        return result[:5]

    def get_counters_for_hero(self, hero_name):
        """
        Returns heroes that have a high Win Rate specifically against the input hero.
        """
        counters = collections.defaultdict(lambda: {'wins': 0, 'total': 0})
        
        # We look for keys where hero_name is the second element (The Opponent)
        # i.e. (Counter_Candidate, My_Hero)
        for (h1, h2), stats in self.counter_map.items():
            if h2 == hero_name:
                counters[h1]['wins'] += stats['wins']
                counters[h1]['total'] += stats['total']
        
        result = []
        for counter, stats in counters.items():
            # Filter to ensure we have enough data (min 2 games)
            if stats['total'] >= 2:
                wr = (stats['wins'] / stats['total']) * 100
                # Only add if they actually have a positive win rate (>50%)
                if wr > 50:
                    result.append({'hero': counter, 'win_rate': wr, 'matches': stats['total']})
        
        result.sort(key=lambda x: (x['win_rate'], x['matches']), reverse=True)
        return result[:5]


    def get_team_signatures(self, team_name):
        """
        Returns detailed insights: Top Pairings, Phase 1 Priority, Phase 2 Priority.
        """
        stats = collections.defaultdict(lambda: {'wins': 0, 'total': 0, 'p1': 0, 'p2': 0})
        pairings = collections.Counter() # Key: tuple(sorted(heroA, heroB)), Value: count
        
        match_count = 0
        
        for match in self.matches:
            is_blue = match['blue_team'] == team_name
            is_red = match['red_team'] == team_name
            
            if not is_blue and not is_red: continue
            match_count += 1

            # Get team data
            if is_blue:
                picks_p1 = match['blue_picks']['p1']
                picks_p2 = match['blue_picks']['p2']
                result = match['blue_result']
            else:
                picks_p1 = match['red_picks']['p1']
                picks_p2 = match['red_picks']['p2']
                result = match['red_result']

            all_picks = [h for h in picks_p1 + picks_p2 if h]

            # Update Stats (WR, Total, Phases)
            for hero in all_picks:
                stats[hero]['total'] += 1
                if result == 'WIN': stats[hero]['wins'] += 1
            
            for hero in picks_p1:
                if hero: stats[hero]['p1'] += 1
            
            for hero in picks_p2:
                if hero: stats[hero]['p2'] += 1

            # Calculate Pairings (Synergy/Combo Frequency)
            # We look at pairs of heroes appearing together
            for pair in itertools.combinations(all_picks, 2):
                key = tuple(sorted(pair))
                pairings[key] += 1
        
        # 1. Process Top Pairings
        top_pairings = []
        for (h1, h2), count in pairings.most_common(10):
            if count >= 2: # Only show if played together at least twice
                top_pairings.append({
                    'heroes': [h1, h2],
                    'count': count
                })

        # 2. Process Priority Picks by Phase
        # Sort by frequency for each phase
        p1_sorted = sorted(stats.items(), key=lambda x: x[1]['p1'], reverse=True)
        p2_sorted = sorted(stats.items(), key=lambda x: x[1]['p2'], reverse=True)
        
        # Get Top 3 for P1 and Top 3 for P2
        priority_p1 = []
        for hero, data in p1_sorted[:3]:
            if data['p1'] > 0:
                wr = (data['wins'] / data['total'] * 100) if data['total'] > 0 else 0
                priority_p1.append({
                    'hero': hero,
                    'count': data['p1'],
                    'win_rate': round(wr, 1)
                })

        priority_p2 = []
        for hero, data in p2_sorted[:3]:
            if data['p2'] > 0:
                wr = (data['wins'] / data['total'] * 100) if data['total'] > 0 else 0
                priority_p2.append({
                    'hero': hero,
                    'count': data['p2'],
                    'win_rate': round(wr, 1)
                })

        return {
            'pairings': top_pairings,
            'priority_p1': priority_p1,
            'priority_p2': priority_p2
        }

    def get_counters_for_team(self, hero_list):
        """
        Aggregates counters for a list of heroes (a team).
        Returns a ranked list of heroes that perform well against the input team.
        """
        threat_map = collections.defaultdict(lambda: {'wins': 0, 'total': 0})
        
        # Iterate through every hero in the user's team to see who beats them
        for my_hero in hero_list:
            if not my_hero: continue
            
            # counter_map Key=(Winner, Loser). 
            # We look for entries where my_hero is the Loser.
            for (winner, loser), stats in self.counter_map.items():
                if loser == my_hero:
                    threat_map[winner]['wins'] += stats['wins']
                    threat_map[winner]['total'] += stats['total']
        
        results = []
        for threat, stats in threat_map.items():
            if stats['total'] >= 2:
                wr = (stats['wins'] / stats['total']) * 100
                # Only include if they have a winning record (>50%)
                if wr > 50.0:
                    results.append({
                        'hero': threat, 
                        'win_rate': round(wr, 1), 
                        'matches': stats['total']
                    })
        
        # Sort by Win Rate desc, then matches desc
        results.sort(key=lambda x: (x['win_rate'], x['matches']), reverse=True)
        return results

    def get_global_meta_stats(self):
        """
        Aggregates overall Pick/Ban/Win rates across all data for a summary.
        """
        hero_stats = collections.defaultdict(lambda: {'picks': 0, 'bans': 0, 'wins': 0, 'games': 0})
        
        for match in self.matches:
            # Count Blue Side
            for h in match['blue_picks']['p1'] + match['blue_picks']['p2']:
                if h: 
                    hero_stats[h]['picks'] += 1
                    hero_stats[h]['games'] += 1
                    if match['blue_result'] == 'WIN': hero_stats[h]['wins'] += 1
            
            for h in match['blue_bans']['p1'] + match['blue_bans']['p2']:
                if h: hero_stats[h]['bans'] += 1

            # Count Red Side
            for h in match['red_picks']['p1'] + match['red_picks']['p2']:
                if h: 
                    hero_stats[h]['picks'] += 1
                    hero_stats[h]['games'] += 1
                    if match['red_result'] == 'WIN': hero_stats[h]['wins'] += 1
            
            for h in match['red_bans']['p1'] + match['red_bans']['p2']:
                if h: hero_stats[h]['bans'] += 1
        
        summary = []
        for hero, stats in hero_stats.items():
            if stats['picks'] > 0:
                wr = (stats['wins'] / stats['picks']) * 100
                pb_rate = stats['picks'] + stats['bans']
                summary.append({
                    'hero': hero,
                    'win_rate': round(wr, 1),
                    'picks': stats['picks'],
                    'bans': stats['bans'],
                    'presence': pb_rate
                })
        
        return summary