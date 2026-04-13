import pandas as pd
import collections
import os
import itertools
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class DataAnalyzer:
    def __init__(self):
        self.matches = [] 
        self.hero_win_rates = collections.defaultdict(lambda: {'wins': 0, 'total': 0})
        self.synergy_map = collections.defaultdict(lambda: {'wins': 0, 'total': 0})
        self.counter_map = collections.defaultdict(lambda: {'wins': 0, 'total': 0})
        
        self.hero_data_dict = None
        self.is_loaded = False
        self._avg_winning_stats_cache = None
        self._load_errors = []

        # --- Performance Optimization: Data Structures for Pandas Analysis ---
        # Instead of just a nested list, we maintain flat records for fast vectorized operations
        self._flat_records = []
        self._matches_df = None
        
        # --- Performance Optimization: Caches for Hero Data ---
        self._hero_lookup = {}  # Maps lowercased names to data (O(1) lookup)
        self._stat_cache = {}   # Caches substring search results
        
        self.load_local_data()

    def get_load_status(self):
        """Return detailed status of data loading for debugging."""
        return {
            "is_loaded": self.is_loaded,
            "match_count": len(self.matches),
            "hero_data_loaded": self.hero_data_dict is not None,
            "hero_data_count": len(self.hero_data_dict) if self.hero_data_dict else 0,
            "load_errors": self._load_errors,
            "available_heroes_in_data": self._get_available_heroes_from_matches()
        }

    def _get_available_heroes_from_matches(self):
        """Get unique hero names from all match picks."""
        heroes = set()
        for match in self.matches:
            for side in ['blue', 'red']:
                picks = match.get(f'{side}_picks', {})
                if picks:
                    for phase in ['p1', 'p2']:
                        for h in picks.get(phase, []):
                            if h and str(h).strip():
                                heroes.add(str(h).strip())
        return sorted(list(heroes))

    def get_hero_name_mismatches(self):
        """Find heroes in matches that don't exist in hero_data."""
        if not self.hero_data_dict:
            return {"error": "Hero data not loaded"}
        
        match_heroes = set(self._get_available_heroes_from_matches())
        data_heroes = set(self.hero_data_dict.keys())
        mismatches = []
        data_heroes_lower = {h.lower(): h for h in data_heroes}
        
        for match_hero in match_heroes:
            if match_hero not in data_heroes:
                if match_hero.lower() in data_heroes_lower:
                    mismatches.append({
                        "match_name": match_hero,
                        "data_name": data_heroes_lower[match_hero.lower()],
                        "type": "case_mismatch"
                    })
                else:
                    mismatches.append({
                        "match_name": match_hero,
                        "data_name": None,
                        "type": "missing_from_data"
                    })
        return mismatches

    def load_local_data(self):
        script_dir = os.path.dirname(os.path.abspath(__file__))
        file_path = os.path.join(script_dir, 'Analyst.xlsx')
        
        if os.path.exists(file_path):
            logger.info(f"Found {file_path}. Processing tournament data...")
            success, msg = self.process_file(file_path)
            if success:
                logger.info(f"Tournament data loaded successfully. {len(self.matches)} matches.")
                self.is_loaded = True
                
                # Optimization: Create Pandas DataFrame after loading all rows
                if self._flat_records:
                    self._matches_df = pd.DataFrame(self._flat_records)
                    logger.info(f"Optimized DataFrame created with {len(self._matches_df)} hero records.")
                
                self._calculate_relationships()
                logger.info("Hero relationships (Synergy/Counters) calculated.")
            else:
                self._load_errors.append(f"Tournament data: {msg}")
                logger.error(f"Failed to load tournament data: {msg}")
        else:
            error_msg = f"File not found at {file_path}"
            self._load_errors.append(error_msg)
            logger.critical(error_msg)

    def process_file(self, file_path):
        try:
            df_raw = pd.read_excel(file_path, sheet_name='MLBB Statistics', header=None)
            
            header_row_idx = None
            # Bug Fix: Case-insensitive header search
            for i, row in df_raw.iterrows():
                row_vals_lower = [str(v).lower().strip() for v in row.values]
                if 'ban 1' in row_vals_lower:
                    header_row_idx = i
                    break
            
            if header_row_idx is None:
                return False, "Header row not found."

            start_row_idx = None
            for i in range(header_row_idx + 1, len(df_raw)):
                val = str(df_raw.iloc[i, 0]).strip() 
                if val and val != '*' and val != 'nan' and 'Tournament' not in val:
                    start_row_idx = i
                    break
            
            if start_row_idx is None:
                return False, "Data rows not found."

            df = df_raw.iloc[start_row_idx:].reset_index(drop=True)
            
            header_row = df_raw.iloc[header_row_idx]
            header_row_norm = header_row.astype(str).str.strip().str.lower()
            ban1_cols = header_row_norm[header_row_norm == 'ban 1'].index.tolist()
            
            if len(ban1_cols) < 2:
                return False, "Column mapping failed."

            col_blue_start = ban1_cols[0]
            col_red_start = ban1_cols[1]
            col_blue_team = col_blue_start - 1
            col_red_team = col_red_start - 1
            
            error_count = 0
            for index, row in df.iterrows():
                try:
                    self._parse_match_row(row, col_blue_team, col_red_team, col_blue_start, col_red_start, index)
                except Exception as e:
                    error_count += 1
                    logger.debug(f"Error parsing row {index}: {e}")
                    continue
            
            if error_count > 0:
                logger.warning(f"Skipped {error_count} rows due to parsing errors")
                
            return True, "Success"
            
        except Exception as e:
            return False, str(e)

    def _parse_match_row(self, row, col_blue_team, col_red_team, col_blue_start, col_red_start, match_id):
        def clean_name(name):
            if pd.isna(name):
                return ""
            return str(name).strip()

        tournament = str(row.iloc[0]).strip()
        selected_map = str(row.iloc[1]).strip()
        blue_team = clean_name(row.iloc[col_blue_team])
        red_team = clean_name(row.iloc[col_red_team])
        
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

        match_data = {
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
        }
        self.matches.append(match_data)
        
        # Optimization: Build flat records for Pandas
        # This allows us to avoid O(N*M) loops later
        def add_records(side, team, picks, bans, result):
            for phase, p_list in picks.items():
                for hero in p_list:
                    if hero:
                        self._flat_records.append({
                            'match_id': match_id,
                            'tournament': tournament,
                            'map': selected_map,
                            'side': side,
                            'team': team,
                            'phase': phase,
                            'type': 'pick',
                            'hero': hero,
                            'result': result
                        })
            for phase, b_list in bans.items():
                for hero in b_list:
                    if hero:
                        self._flat_records.append({
                            'match_id': match_id,
                            'tournament': tournament,
                            'map': selected_map,
                            'side': side,
                            'team': team,
                            'phase': phase,
                            'type': 'ban',
                            'hero': hero,
                            'result': None
                        })

        add_records('blue', blue_team, match_data['blue_picks'], match_data['blue_bans'], blue_result)
        add_records('red', red_team, match_data['red_picks'], match_data['red_bans'], red_result)

        all_blue = [p for p in blue_picks_p1 + blue_picks_p2 if p]
        all_red = [p for p in red_picks_p1 + red_picks_p2 if p]
        
        self._update_hero_stats(all_blue, blue_result)
        self._update_hero_stats(all_red, red_result)

    def _update_hero_stats(self, picks, result):
        for hero in picks:
            self.hero_win_rates[hero]['total'] += 1
            if result == 'WIN':
                self.hero_win_rates[hero]['wins'] += 1

    def _calculate_relationships(self):
        logger.info("Calculating Hero Relationships...")
        for match in self.matches:
            blue_picks = [h for h in match['blue_picks']['p1'] + match['blue_picks']['p2'] if h]
            red_picks = [h for h in match['red_picks']['p1'] + match['red_picks']['p2'] if h]
            blue_win = match['blue_result'] == 'WIN'
            
            winning_team = blue_picks if blue_win else red_picks
            losing_team = red_picks if blue_win else blue_picks
            
            for pair in itertools.combinations(winning_team, 2):
                key = tuple(sorted(pair))
                self.synergy_map[key]['total'] += 1
                self.synergy_map[key]['wins'] += 1
            for pair in itertools.combinations(losing_team, 2):
                key = tuple(sorted(pair))
                self.synergy_map[key]['total'] += 1
            
            for winner_hero in winning_team:
                for loser_hero in losing_team:
                    key_win = (winner_hero, loser_hero)
                    self.counter_map[key_win]['wins'] += 1
                    self.counter_map[key_win]['total'] += 1
                    
                    key_loss = (loser_hero, winner_hero)
                    self.counter_map[key_loss]['total'] += 1

    def get_unique_values(self, column, tournament_filter=None):
        if not self.matches:
            return []
        # Optimization: Use DataFrame if available
        if self._matches_df is not None:
            if column == 'team':
                return sorted(self._matches_df['team'].unique().tolist())
            elif column in ['tournament', 'map']:
                return sorted(self._matches_df[column].unique().tolist())
        
        # Fallback
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
        """Optimized using Pandas DataFrame for vectorized aggregation."""
        if self._matches_df is None or self._matches_df.empty:
            return pd.DataFrame(columns=['Hero', 'Pick P1', 'Pick P2', 'Ban P1', 'Ban P2', 'Total Picks', 'Win Rate'])
            
        df = self._matches_df.copy()
        
        # Apply Filters
        if side:
            df = df[df['side'] == side.lower()]
        if tournament_filter and tournament_filter != "All":
            df = df[df['tournament'] == tournament_filter]
        if map_filter and map_filter != "All":
            df = df[df['map'] == map_filter]
        if team_filter and team_filter != "All":
            df = df[df['team'] == team_filter]
            
        # Separate Picks and Bans
        picks_df = df[df['type'] == 'pick']
        bans_df = df[df['type'] == 'ban']
        
        # Helper to aggregate counts
        def get_counts(df_source):
            return df_source.groupby('hero').size()

        p1_picks = get_counts(picks_df[picks_df['phase'] == 'p1'])
        p2_picks = get_counts(picks_df[picks_df['phase'] == 'p2'])
        p1_bans = get_counts(bans_df[bans_df['phase'] == 'p1'])
        p2_bans = get_counts(bans_df[bans_df['phase'] == 'p2'])
        
        total_picks = get_counts(picks_df)
        wins = get_counts(picks_df[picks_df['result'] == 'WIN'])
        
        # Build Result DataFrame
        stats = pd.DataFrame({
            'Pick P1': p1_picks,
            'Pick P2': p2_picks,
            'Ban P1': p1_bans,
            'Ban P2': p2_bans,
            'Total Picks': total_picks,
            'wins': wins
        }).fillna(0).astype(int)
        
        # Calculate Win Rate safely
        stats['Win Rate'] = (stats['wins'] / stats['Total Picks'].replace(0, 1) * 100).round(2)
        stats.loc[stats['Total Picks'] == 0, 'Win Rate'] = 0.0
        
        stats = stats.drop(columns=['wins'])
        stats.index.name = 'Hero'
        stats = stats.reset_index()
        
        return stats.sort_values(by='Total Picks', ascending=False)

    def get_hero_win_rate(self, hero_name):
        data = self.hero_win_rates.get(hero_name)
        if not data or data['total'] == 0:
            return None
        return (data['wins'] / data['total']) * 100

    def get_synergy_for_hero(self, hero_name):
        partners = collections.defaultdict(lambda: {'wins': 0, 'total': 0})
        for (h1, h2), stats in self.synergy_map.items():
            if h1 == hero_name:
                partners[h2]['wins'] += stats['wins']
                partners[h2]['total'] += stats['total']
            elif h2 == hero_name:
                partners[h1]['wins'] += stats['wins']
                partners[h1]['total'] += stats['total']
        result = []
        for partner, stats in partners.items():
            if stats['total'] >= 2:
                wr = (stats['wins'] / stats['total']) * 100
                result.append({'hero': partner, 'win_rate': wr, 'matches': stats['total']})
        result.sort(key=lambda x: (x['matches'], x['win_rate']), reverse=True)
        return result[:5]

    def get_counters_for_hero(self, hero_name):
        counters = collections.defaultdict(lambda: {'wins': 0, 'total': 0})
        for (h1, h2), stats in self.counter_map.items():
            if h2 == hero_name:
                counters[h1]['wins'] += stats['wins']
                counters[h1]['total'] += stats['total']
        result = []
        for counter, stats in counters.items():
            if stats['total'] >= 2:
                wr = (stats['wins'] / stats['total']) * 100
                if wr > 50:
                    result.append({'hero': counter, 'win_rate': wr, 'matches': stats['total']})
        result.sort(key=lambda x: (x['matches'], x['win_rate']), reverse=True)
        return result[:5]

    def get_team_signatures(self, team_name, tournament_filter=None):
        stats_blue = {} 
        stats_red = {} 
        pairings = {}
        matches_played = 0
        
        def init_hero(dict_ref, h):
            dict_ref[h] = {'wins': 0, 'total': 0, 'picks': 0}

        for match in self.matches:
            # 1. Filter by Tournament
            match_tourn = match.get('tournament')
            if tournament_filter and tournament_filter != "All" and match_tourn != tournament_filter:
                continue
                
            # 2. Check if team is in match
            blue_team = match.get('blue_team')
            red_team = match.get('red_team')
            
            if not blue_team or not red_team: continue

            is_blue = (blue_team == team_name)
            is_red = (red_team == team_name)
            
            if not is_blue and not is_red:
                continue
            
            matches_played += 1
            
            # 3. Extract Data Safely
            try:
                side = 'blue' if is_blue else 'red'
                
                picks_data = match.get(f'{side}_picks')
                if not isinstance(picks_data, dict): continue 
                
                # We combine P1 and P2 for the side analysis
                picks_p1 = picks_data.get('p1', [])
                picks_p2 = picks_data.get('p2', [])
                
                result = match.get(f'{side}_result')
                if result: result = str(result).upper()
                
                if not isinstance(picks_p1, list): picks_p1 = []
                if not isinstance(picks_p2, list): picks_p2 = []
                
                all_picks = [h for h in picks_p1 + picks_p2 if h and str(h).strip()]
                
                # --- UPDATE STATS BASED ON SIDE ---
                target_stats = stats_blue if is_blue else stats_red
                
                for hero in all_picks:
                    if hero not in target_stats: init_hero(target_stats, hero)
                    target_stats[hero]['picks'] += 1
                    target_stats[hero]['total'] += 1
                    if result == 'WIN':
                        target_stats[hero]['wins'] += 1
                # ----------------------------------

                # --- UPDATE GLOBAL PAIRINGS (Combined) ---
                for pair in itertools.combinations(all_picks, 2):
                    h1, h2 = sorted(pair)
                    key = (h1, h2)
                    if key not in pairings: pairings[key] = 0
                    pairings[key] += 1
                # --------------------------------------

            except Exception:
                continue
        
        # 4. Build Result
        if matches_played == 0:
            return {
                'pairings': [], 
                'priority_blue_side': [], 
                'priority_red_side': [],
                'matches_played': 0,
                'found': False
            }
        
        # --- HELPER TO GENERATE LISTS ---
        def get_priority_list(source_stats):
            temp_list = []
            for hero, s_data in source_stats.items():
                count = s_data.get('picks', 0)
                if count > 0:
                    wins = s_data.get('wins', 0)
                    total = s_data.get('total', 1)
                    win_rate = (wins / total * 100) if total > 0 else 0
                    
                    temp_list.append({
                        'hero': hero, 
                        'count': count, 
                        'win_rate': round(win_rate, 1)
                    })
            
            temp_list.sort(key=lambda x: (x['count'], x['win_rate']), reverse=True)
            return temp_list[:5]
        
        sorted_pairings = sorted(pairings.items(), key=lambda x: x[1], reverse=True)
        top_pairings = [{'heroes': list(pair), 'count': count} for pair, count in sorted_pairings if count >= 2][:20]

        return {
            'pairings': top_pairings, 
            'priority_blue_side': get_priority_list(stats_blue), 
            'priority_red_side': get_priority_list(stats_red),
            'matches_played': matches_played,
            'found': True
        }

    def _is_team_in_match(self, match, team_name):
        return match['blue_team'] == team_name or match['red_team'] == team_name

    def _update_signature_stats(self, stats, all_picks, picks_p1, picks_p2, result):
        for hero in all_picks:
            stats[hero]['total'] += 1
            if result == 'WIN':
                stats[hero]['wins'] += 1
        for hero in picks_p1:
            if hero: stats[hero]['p1'] += 1
        for hero in picks_p2:
            if hero: stats[hero]['p2'] += 1

    def _update_pairings(self, pairings, all_picks):
        for pair in itertools.combinations(all_picks, 2):
            key = tuple(sorted(pair))
            pairings[key] += 1

    def _build_team_signatures(self, stats, pairings):
        top_pairings = [{'heroes': list(pair), 'count': count} for pair, count in pairings.most_common(10) if count >= 2]
        priority_p1 = self._get_priority_picks(stats, 'p1')[:3]
        priority_p2 = self._get_priority_picks(stats, 'p2')[:3]
        return {'pairings': top_pairings, 'priority_p1': priority_p1, 'priority_p2': priority_p2}

    def _get_priority_picks(self, stats, phase):
        result = []
        sorted_stats = sorted(stats.items(), key=lambda x: x[1][phase], reverse=True)
        for hero, data in sorted_stats:
            if data[phase] > 0:
                wr = (data['wins'] / data['total'] * 100) if data['total'] > 0 else 0
                result.append({'hero': hero, 'count': data[phase], 'win_rate': round(wr, 1)})
        return result

    def get_counters_for_team(self, hero_list):
        threat_map = collections.defaultdict(lambda: {'wins': 0, 'total': 0})
        for my_hero in hero_list:
            if not my_hero: continue
            for (winner, loser), stats in self.counter_map.items():
                if loser == my_hero:
                    threat_map[winner]['wins'] += stats['wins']
                    threat_map[winner]['total'] += stats['total']
        results = []
        for threat, stats in threat_map.items():
            if stats['total'] >= 2:
                wr = (stats['wins'] / stats['total']) * 100
                if wr > 50.0:
                    results.append({'hero': threat, 'win_rate': round(wr, 1), 'matches': stats['total']})
        results.sort(key=lambda x: (x['win_rate'], x['matches']), reverse=True)
        return results

    def get_global_meta_stats(self):
        """Optimized using Pandas."""
        if self._matches_df is None or self._matches_df.empty:
            return []

        picks_df = self._matches_df[self._matches_df['type'] == 'pick']
        bans_df = self._matches_df[self._matches_df['type'] == 'ban']
        wins_df = picks_df[picks_df['result'] == 'WIN']

        picks_count = picks_df.groupby('hero').size()
        bans_count = bans_df.groupby('hero').size()
        wins_count = wins_df.groupby('hero').size()

        meta = pd.DataFrame({
            'picks': picks_count,
            'bans': bans_count,
            'wins': wins_count
        }).fillna(0).astype(int)
        
        meta['presence'] = meta['picks'] + meta['bans']
        meta['win_rate'] = (meta['wins'] / meta['picks'].replace(0, 1) * 100).round(1)
        meta.loc[meta['picks'] == 0, 'win_rate'] = 0.0

        result = []
        for hero, row in meta.iterrows():
            if row['picks'] > 0:
                result.append({
                    'hero': hero,
                    'win_rate': row['win_rate'],
                    'picks': int(row['picks']),
                    'bans': int(row['bans']),
                    'presence': int(row['presence'])
                })
        
        return sorted(result, key=lambda x: x['picks'], reverse=True)

    def set_hero_data(self, hero_data):
        self.hero_data_dict = hero_data
        # Optimization: Build case-insensitive lookup map
        self._hero_lookup = {k.lower(): v for k, v in hero_data.items()}
        self._avg_winning_stats_cache = None
        self._stat_cache = {} 

    def _get_hero_stat(self, hero_name, stat_name):
        if not hero_name or not self.hero_data_dict:
            return 0.0
        
        key = hero_name.lower()
        
        # Optimization 1: O(1) Lookup
        if key in self._hero_lookup:
            return self._hero_lookup[key].get('stats', {}).get(stat_name, 0.0)
        
        # Optimization 2: Fallback substring search (cached)
        cache_key = (hero_name, stat_name)
        if cache_key in self._stat_cache:
            return self._stat_cache[cache_key]
            
        val = 0.0
        for dict_key, data in self.hero_data_dict.items():
            if key in dict_key.lower() or dict_key.lower() in key:
                val = data.get('stats', {}).get(stat_name, 0.0)
                break
        
        self._stat_cache[cache_key] = val
        return val

    def _calculate_team_stats(self, team_heroes):
        stats = {'Durability': 0, 'Offense': 0, 'Crowd Control': 0, 'Mobility': 0, 'Wave Control': 0}
        valid_heroes = 0
        for hero in team_heroes:
            if self._get_hero_stat(hero, 'Durability') > 0:
                valid_heroes += 1
            for stat_name in stats:
                stats[stat_name] += self._get_hero_stat(hero, stat_name)
        return stats, valid_heroes

    def _classify_archetype(self, team_heroes):
        if not team_heroes:
            return {"Split Push": 0, "Team Fight": 0, "Pick Off": 0}
        
        stats, valid_count = self._calculate_team_stats(team_heroes)
        
        if valid_count == 0:
            return {"Split Push": 0, "Team Fight": 0, "Pick Off": 0}
        
        scale_factor = 5.0 / valid_count
        
        wc = stats['Wave Control'] * scale_factor
        mob = stats['Mobility'] * scale_factor
        dur = stats['Durability'] * scale_factor
        cc = stats['Crowd Control'] * scale_factor
        off = stats['Offense'] * scale_factor
        
        split_push_numerator = (3.0 * wc) + (2.0 * off) + (1.0 * mob) + (0.5 * dur)
        split_push = split_push_numerator / 6.5

        team_fight_numerator = (2.0 * cc) + (2.5 * dur) + (2.0 * off) + (1.0 * mob)
        team_fight = team_fight_numerator / 7.5

        pick_off_numerator = (3.0 * off) + (2.0 * mob) + (1.0 * cc)
        pick_off = pick_off_numerator / 6.0
        
        return {
            "Split Push": round(split_push, 1),
            "Team Fight": round(team_fight, 1),
            "Pick Off": round(pick_off, 1)
        }

    def get_archetype_stats(self):
        arch_stats = collections.defaultdict(lambda: {'wins': 0, 'total': 0})
        for match in self.matches:
            for side in ['blue', 'red']:
                picks = [h for h in match[f'{side}_picks']['p1'] + match[f'{side}_picks']['p2'] if h]
                if len(picks) >= 3:
                    potentials = self._classify_archetype(picks)
                    primary_arch = max(potentials, key=potentials.get)
                    
                    arch_stats[primary_arch]['total'] += 1
                    if match[f'{side}_result'] == 'WIN':
                        arch_stats[primary_arch]['wins'] += 1
        return sorted(
            [{'archetype': arch, 'wins': data['wins'], 'total': data['total'],
              'win_rate': round((data['wins'] / data['total']) * 100, 1)}
             for arch, data in arch_stats.items() if data['total'] > 0],
            key=lambda x: x['win_rate'], reverse=True
        )

    def _get_avg_winning_stats(self):
        if self._avg_winning_stats_cache is not None:
            return self._avg_winning_stats_cache
        
        winning_stats = {'Durability': [], 'Offense': [], 'Crowd Control': [], 'Mobility': [], 'Wave Control': []}
        for match in self.matches:
            winner_picks = []
            if match['blue_result'] == 'WIN':
                winner_picks = [h for h in match['blue_picks']['p1'] + match['blue_picks']['p2'] if h]
            else:
                winner_picks = [h for h in match['red_picks']['p1'] + match['red_picks']['p2'] if h]
            
            if len(winner_picks) >= 3:
                stats, valid_count = self._calculate_team_stats(winner_picks)
                if valid_count > 0:
                    scale = 5.0 / valid_count
                    for k in stats:
                        winning_stats[k].append(stats[k] * scale)
        
        count = len(winning_stats['Durability'])
        if count > 0:
            self._avg_winning_stats_cache = {k: sum(v) / count for k, v in winning_stats.items()}
        else:
            self._avg_winning_stats_cache = {'Durability': 30.0, 'Offense': 32.0, 'Crowd Control': 25.0, 'Mobility': 28.0, 'Wave Control': 22.0}
            logger.warning("No winning team data found, using default averages")
        return self._avg_winning_stats_cache

    def analyze_comp_archetype(self, blue_team, red_team):
        blue_stats, blue_valid = self._calculate_team_stats(blue_team) if blue_team else ({}, 0)
        red_stats, red_valid = self._calculate_team_stats(red_team) if red_team else ({}, 0)
        
        blue_potentials = self._classify_archetype(blue_team) if blue_team else {"Split Push": 0, "Team Fight": 0, "Pick Off": 0}
        red_potentials = self._classify_archetype(red_team) if red_team else {"Split Push": 0, "Team Fight": 0, "Pick Off": 0}
        
        blue_normalized = self._normalize_stats(blue_stats, len(blue_team)) if blue_team else None
        red_normalized = self._normalize_stats(red_stats, len(red_team)) if red_team else None
        
        blue_strat = self._get_strategy(blue_potentials, red_potentials)
        red_strat = self._get_strategy(red_potentials, blue_potentials)
        
        return {
            'blue_stats': blue_normalized, 
            'red_stats': red_normalized,
            'blue_potentials': blue_potentials,
            'red_potentials': red_potentials,
            'blue_strategy': blue_strat,
            'red_strategy': red_strat,
            'blue_hero_count': len(blue_team), 
            'red_hero_count': len(red_team),
            'is_complete': len(blue_team) == 5 and len(red_team) == 5
        }

    def _get_strategy(self, my_potentials, enemy_potentials):
        if not my_potentials or not enemy_potentials:
            return "Calculating..."
        
        strategy_map = {
            "Split Push":
            "Utilize one heroes to pressure a side lane, creating isolated 1v1 situations and forcing the enemy to respond, which can open opportunities for your team elsewhere on the map.",
            "Team Fight":
            "Coordinate as a group to engage in battles around key neutral objectives in the river, leveraging team synergy and positioning to secure victories.",
            "Pick Off":
            "Establish vision control in jungle areas to ambush and eliminate isolated enemy heroes, creating a numbers advantage for your team."
        }
        
        sorted_mine = sorted(my_potentials.items(), key=lambda x: x[1], reverse=True)
        
        for arch_name, my_score in sorted_mine:
            enemy_score = enemy_potentials.get(arch_name, 0)
            if my_score > enemy_score:
                return strategy_map.get(arch_name, "Play standard macro.")
        
        min_diff = float('inf')
        closest_arch = None
        
        for arch_name, my_score in sorted_mine:
            enemy_score = enemy_potentials.get(arch_name, 0)
            diff = abs(enemy_score - my_score)
            if diff < min_diff:
                min_diff = diff
                closest_arch = arch_name
                
        if closest_arch:
            return strategy_map.get(closest_arch, "Play standard macro.")
        return "Play standard macro."
    
    def _normalize_stats(self, stats, hero_count):
        if not stats or hero_count == 0: return None
        if hero_count == 5: return stats
        scale = 5.0 / hero_count
        return {k: v * scale for k, v in stats.items()}

    def get_map_win_rates(self):
        if self._matches_df is None or self._matches_df.empty:
            return []
            
        match_map_counts = self._matches_df[['match_id', 'map']].drop_duplicates().groupby('map').size()
        
        blue_wins_df = self._matches_df[(self._matches_df['side'] == 'blue') & (self._matches_df['result'] == 'WIN')]
        blue_wins_counts = blue_wins_df[['match_id', 'map']].drop_duplicates().groupby('map').size()
        
        result = []
        for map_name, total in match_map_counts.items():
            blue_w = blue_wins_counts.get(map_name, 0)
            red_w = total - blue_w 
            blue_wr = (blue_w / total) * 100
            red_wr = (red_w / total) * 100
            
            result.append({
                'map': map_name,
                'total_matches': int(total),
                'blue_win_rate': round(blue_wr, 1),
                'red_win_rate': round(red_wr, 1),
                'dominant_side': 'Blue' if blue_wr > 50 else 'Red' if red_wr > 50 else 'Balanced'
            })
        
        return sorted(result, key=lambda x: x['total_matches'], reverse=True)

    def get_map_leaders(self, map_name):
        """Optimized using Pandas."""
        if self._matches_df is None or self._matches_df.empty:
            return []
            
        map_df = self._matches_df[self._matches_df['map'] == map_name]
        picks_df = map_df[map_df['type'] == 'pick']
        
        if picks_df.empty:
            return []
            
        hero_stats = picks_df.groupby('hero').agg(
            picks=('hero', 'size'),
            wins=('result', lambda x: (x == 'WIN').sum())
        ).reset_index()
        
        hero_stats['win_rate'] = (hero_stats['wins'] / hero_stats['picks'] * 100).round(1)
        
        leaders = hero_stats[hero_stats['picks'] >= 2].sort_values(
            by=['picks', 'win_rate'], ascending=[False, False]
        )
        
        return leaders.head(5).to_dict('records')
    
    def get_suggested_bans(self, side='Blue', tournament_filter=None):
        """
        Returns top ban targets based on tournament statistics for a specific side (Blue/Red).
        Metrics: Ban Frequency and Ban Win Rate (Win rate when hero was banned).
        """
        if self._matches_df is None or self._matches_df.empty:
            return []
        
        df = self._matches_df.copy()
        
        # 1. Filter by Side
        df = df[df['side'] == side.lower()]
        
        # 2. Filter by Tournament
        if tournament_filter and tournament_filter != "All":
            df = df[df['tournament'] == tournament_filter]
            
        if df.empty:
            return []
            
        # 3. Identify Winning Matches for this side
        winning_match_ids = df[(df['type'] == 'pick') & (df['result'] == 'WIN')]['match_id'].unique()
        
        # 4. Filter Bans
        df_bans = df[df['type'] == 'ban']
        
        if df_bans.empty:
            return []
        
        # 5. Total Ban Count per Hero
        total_bans = df_bans.groupby('hero').size().reset_index(name='total_bans')
        
        # 6. Winning Ban Count (Effective Bans)
        winning_bans = df_bans[df_bans['match_id'].isin(winning_match_ids)]
        winning_ban_counts = winning_bans.groupby('hero').size().reset_index(name='winning_bans')
        
        # 7. Merge and Calculate Stats
        stats = total_bans.merge(winning_ban_counts, on='hero', how='left').fillna(0)
        stats['winning_bans'] = stats['winning_bans'].astype(int)
        
        # Ban Win Rate: (Matches Won with Ban) / (Total Matches with Ban)
        stats['ban_win_rate'] = (stats['winning_bans'] / stats['total_bans'] * 100).round(1)
        
        # Total matches played by this side to calculate Ban Frequency percentage
        total_matches_played = df[['match_id']].drop_duplicates().shape[0]
        
        stats['ban_frequency'] = (stats['total_bans'] / total_matches_played * 100).round(1)
        
        # Sort: First by Ban Frequency (Meta), then by Ban Win Rate (Effectiveness)
        stats = stats.sort_values(by=['total_bans', 'ban_win_rate'], ascending=[False, False])
        
        return stats.head(5).to_dict('records')