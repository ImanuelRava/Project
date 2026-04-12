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
            for i, row in df_raw.iterrows():
                if 'Ban 1' in row.values:
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
                    self._parse_match_row(row, col_blue_team, col_red_team, col_blue_start, col_red_start)
                except Exception as e:
                    error_count += 1
                    logger.debug(f"Error parsing row {index}: {e}")
                    continue
            
            if error_count > 0:
                logger.warning(f"Skipped {error_count} rows due to parsing errors")
                
            return True, "Success"
            
        except Exception as e:
            return False, str(e)

    def _parse_match_row(self, row, col_blue_team, col_red_team, col_blue_start, col_red_start):
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
        """Get hero summary with better error handling."""
        try:
            stats = collections.defaultdict(lambda: {
                'picks_p1': 0, 'picks_p2': 0, 'bans_p1': 0, 'bans_p2': 0, 
                'wins': 0, 'total_picks': 0
            })
            
            for match in self.matches:
                if not self._passes_filters(match, tournament_filter, map_filter):
                    continue
                
                target_picks, target_bans, target_result = self._get_match_side_data(match, side, team_filter)
                if target_picks is None:
                    continue
                
                self._update_side_stats(stats, target_picks, target_bans, target_result)
            
            return self._build_hero_summary_df(stats)
            
        except Exception as e:
            logger.error(f"Error in get_hero_summary: {e}", exc_info=True)
            return pd.DataFrame(columns=['Hero', 'Pick P1', 'Pick P2', 'Ban P1', 'Ban P2', 'Total Picks', 'Win Rate'])

    def _passes_filters(self, match, tournament_filter, map_filter):
        if tournament_filter and tournament_filter != "All" and match['tournament'] != tournament_filter:
            return False
        if map_filter and map_filter != "All" and match['map'] != map_filter:
            return False
        return True

    def _get_match_side_data(self, match, side, team_filter):
        if side == 'Blue':
            if team_filter and team_filter != "All" and match['blue_team'] != team_filter:
                return None, None, None
            return match['blue_picks'], match['blue_bans'], match['blue_result']
        else:
            if team_filter and team_filter != "All" and match['red_team'] != team_filter:
                return None, None, None
            return match['red_picks'], match['red_bans'], match['red_result']

    def _update_side_stats(self, stats, target_picks, target_bans, target_result):
        for phase in ['p1', 'p2']:
            for hero in target_picks.get(phase, []):
                if hero:
                    stats[hero][f'picks_{phase}'] += 1
                    stats[hero]['total_picks'] += 1
                    if target_result == 'WIN':
                        stats[hero]['wins'] += 1
            for hero in target_bans.get(phase, []):
                if hero:
                    stats[hero][f'bans_{phase}'] += 1

    def _build_hero_summary_df(self, stats):
        data = []
        for hero, s in stats.items():
            wr = (s['wins'] / s['total_picks'] * 100) if s['total_picks'] > 0 else 0.0
            data.append({
                'Hero': hero, 'Pick P1': s['picks_p1'], 'Pick P2': s['picks_p2'],
                'Ban P1': s['bans_p1'], 'Ban P2': s['bans_p2'],
                'Total Picks': s['total_picks'], 'Win Rate': wr
            })
        return pd.DataFrame(data)

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
        result.sort(key=lambda x: (x['win_rate'], x['matches']), reverse=True)
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
        result.sort(key=lambda x: (x['win_rate'], x['matches']), reverse=True)
        return result[:5]

    def get_team_signatures(self, team_name):
        stats = collections.defaultdict(lambda: {'wins': 0, 'total': 0, 'p1': 0, 'p2': 0})
        pairings = collections.Counter()
        
        for match in self.matches:
            if not self._is_team_in_match(match, team_name):
                continue
            is_blue = match['blue_team'] == team_name
            picks_p1 = match['blue_picks']['p1'] if is_blue else match['red_picks']['p1']
            picks_p2 = match['blue_picks']['p2'] if is_blue else match['red_picks']['p2']
            result = match['blue_result'] if is_blue else match['red_result']
            all_picks = [h for h in picks_p1 + picks_p2 if h]

            self._update_signature_stats(stats, all_picks, picks_p1, picks_p2, result)
            self._update_pairings(pairings, all_picks)
        
        return self._build_team_signatures(stats, pairings)

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
        hero_stats = collections.defaultdict(lambda: {'picks': 0, 'bans': 0, 'wins': 0})
        for match in self.matches:
            self._update_global_stats(hero_stats, match['blue_picks'], match['blue_bans'], match['blue_result'])
            self._update_global_stats(hero_stats, match['red_picks'], match['red_bans'], match['red_result'])
        return [
            {'hero': hero, 'win_rate': round((s['wins'] / s['picks']) * 100, 1), 
             'picks': s['picks'], 'bans': s['bans'], 'presence': s['picks'] + s['bans']}
            for hero, s in hero_stats.items() if s['picks'] > 0
        ]

    def _update_global_stats(self, hero_stats, picks, bans, result):
        for phase in ['p1', 'p2']:
            for h in picks.get(phase, []):
                if h:
                    hero_stats[h]['picks'] += 1
                    if result == 'WIN': hero_stats[h]['wins'] += 1
            for h in bans.get(phase, []):
                if h: hero_stats[h]['bans'] += 1

    def set_hero_data(self, hero_data):
        self.hero_data_dict = hero_data
        self._avg_winning_stats_cache = None

    def _get_hero_stat(self, hero_name, stat_name):
        if not self.hero_data_dict:
            return 0.0
        if hero_name in self.hero_data_dict:
            return self.hero_data_dict[hero_name]['stats'].get(stat_name, 0.0)
        
        hero_name_lower = hero_name.lower()
        for key, data in self.hero_data_dict.items():
            if key.lower() == hero_name_lower:
                return data['stats'].get(stat_name, 0.0)
        for key, data in self.hero_data_dict.items():
            if hero_name_lower in key.lower() or key.lower() in hero_name_lower:
                return data['stats'].get(stat_name, 0.0)
        return 0.0

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
        
        # 1. Split Push
        # Weights: 3.0, 2.0, 1.0, 0.5 (Total: 6.5)
        split_push_numerator = (3.0 * wc) + (2.0 * off) + (1.0 * mob) + (0.5 * dur)
        split_push = split_push_numerator / 6.5

        # 2. Team Fight
        # Weights: 2.0, 2.5, 2.0, 1.0 (Total: 7.5)
        team_fight_numerator = (2.0 * cc) + (2.5 * dur) + (2.0 * off) + (1.0 * mob)
        team_fight = team_fight_numerator / 7.5

        # 3. Pick Off
        # Weights: 3.0, 2.0, 1.0 (Total: 6.0)
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
        """Determine strategy by finding an archetype where I out-stat the enemy.
        If I lose all matchups, fall back to the archetype closest in value to the enemy team."""
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
        
        # Step 1: Sort my archetypes from highest score to lowest
        sorted_mine = sorted(my_potentials.items(), key=lambda x: x[1], reverse=True)
        
        # Step 2: Go down the list. If my score > enemy's score for that archetype, use it.
        for arch_name, my_score in sorted_mine:
            enemy_score = enemy_potentials.get(arch_name, 0)
            if my_score > enemy_score:
                return strategy_map.get(arch_name, "Play standard macro.")
        
        # Step 3: If I lose all matchups, find the archetype with the closest value to the enemy team
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