import pandas as pd
import random
import os
from flask import Flask, jsonify, request, render_template
from flask_cors import CORS
from data_analyzer import DataAnalyzer

# --- CONFIGURATION ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE = os.path.join(BASE_DIR, 'Analyst.xlsx')
SHEET_NAME = 'Hero Data'

analyzer = DataAnalyzer()
app = Flask(__name__)
CORS(app)

# --- GLOBAL DATA STORE ---
hero_data = {}
available_heroes_list = []

def generate_color(name):
    """Generates a consistent hex color based on the hero name."""
    random.seed(name)
    r = lambda: random.randint(50, 200)
    return '#%02X%02X%02X' % (r(), r(), r())

def load_data():
    """Extracts info from Hero Data sheet and structures it."""
    global hero_data, available_heroes_list

    df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)

    df.replace('N/A', pd.NA, inplace=True)

    stats_columns = ['Durability', 'Offense', 'Crowd Control', 'Mobility', 'Wave Control']

    for index, row in df.iterrows():
        name = row.get('Hero')
        if pd.isna(name): continue

        name = str(name).strip()

        # Extract Roles and Lanes
        roles = [str(row.get('Role 1')).strip()]
        role2 = row.get('Role 2')
        if pd.notna(role2) and str(role2).strip().lower() != 'nan':
            roles.append(str(role2).strip())

        lanes = [str(row.get('Lane 1')).strip()]
        lane2 = row.get('Lane 2')
        if pd.notna(lane2) and str(lane2).strip().lower() != 'nan':
            lanes.append(str(lane2).strip())

        # Extract Stats
        stats = {}
        for stat in stats_columns:
            val = row.get(stat)
            stats[stat] = float(val) if pd.notna(val) else 0.0

        image_path = f"/static/hero_icon/{name}.png"

        hero_data[name] = {
            "roles": roles,
            "lanes": lanes,
            "stats": stats,
            "color": generate_color(name), 
            "image": image_path            
        }
        available_heroes_list.append(name)

    print(f"Loaded {len(hero_data)} heroes from {SHEET_NAME}.")

def calculate_weighted_stats(team_list):
    categories = ['Durability', 'Offense', 'Crowd Control', 'Mobility', 'Wave Control']
    weighted_stats = {cat: 0.0 for cat in categories}

    if not team_list:
        return weighted_stats

    for category in categories:
        hero_values = []
        for hero_name in team_list:
            if hero_name in hero_data:
                val = hero_data[hero_name]['stats'].get(category, 0)
                hero_values.append(val)

        if not hero_values:
            weighted_stats[category] = 0
            continue

        max_val = max(hero_values)
        total_sum = sum(hero_values)
        count_others = len(hero_values) - 1

        if count_others > 0:
            sum_others = total_sum - max_val
            avg_others = sum_others / count_others
        else:
            avg_others = 0

        #Scoring: 40% weight to the highest stat, 60% to the average of the others
        weighted_score = (max_val * 0.4) + (avg_others * 0.6)
        weighted_stats[category] = weighted_score

    return weighted_stats

def validate_team_composition(team_list):
    """
    Returns True if the team adheres to:
    1. Exactly 1 Mage.
    2. Exactly 1 Marksman.
    3. All 5 standard lanes are filled.
    """
    roles_count = {"Mage": 0, "Marksman": 0}
    lanes_covered = set()
    
    # Standard Lanes that MUST be covered
    required_lanes = ["Gold Lane", "EXP Lane", "Mid Lane", "Jungle", "Roaming"]

    for h in team_list:
        if h in hero_data:
            data = hero_data[h]
            
            # Count Roles
            for role in data['roles']:
                if role in roles_count:
                    roles_count[role] += 1
            
            # Track Lanes (Ignore N/A for strict lane filling check)
            for lane in data['lanes']:
                if lane and lane.lower() != 'n/a' and lane.lower() != 'nan':
                    lanes_covered.add(lane)

    # Check Constraints
    has_one_mage = roles_count["Mage"] == 1
    has_one_mm = roles_count["Marksman"] == 1
    has_all_lanes = all(lane in lanes_covered for lane in required_lanes)

    return has_one_mage and has_one_mm and has_all_lanes

def get_meta_popularity():
    """Calculates a popularity score (Total Picks + Total Bans)."""
    popularity = {}
    if not analyzer.matches:
        return popularity

    for match in analyzer.matches:
        all_hero_lists = []
        
        def safe_extend(target_list, source_dict):
            if source_dict:
                p1 = source_dict.get('p1', [])
                p2 = source_dict.get('p2', [])
                valid = [h for h in p1 + p2 if h and str(h).strip() != '']
                target_list.extend(valid)

        safe_extend(all_hero_lists, match.get('blue_picks'))
        safe_extend(all_hero_lists, match.get('red_picks'))
        safe_extend(all_hero_lists, match.get('blue_bans'))
        safe_extend(all_hero_lists, match.get('red_bans'))

        for hero in all_hero_lists:
            popularity[hero] = popularity.get(hero, 0) + 1
            
    return popularity

def evaluate_matchup_score(my_stats, enemy_stats):
    """Returns a win probability (0.0 to 1.0)."""
    my_power = sum(my_stats.values())
    enemy_power = sum(enemy_stats.values())
    total_power = my_power + enemy_power
    
    if total_power == 0:
        return 0.5
    else:
        return my_power / total_power

def get_smart_suggestion(my_team, enemy_team, my_stats, enemy_stats, my_side_name):
    """
    Unified Suggestion Logic:
    - Both ADD and SWAP phases use "Win Rate Simulation" to determine the best option.
    - This ensures consistency: The hero you pick today is the one that stays in the team tomorrow.
    - Strictly enforces: 1 Mage, 1 Marksman, 5 Lanes, Meta Pool.
    """
    picked = set(my_team + enemy_team)
    meta_scores = get_meta_popularity()
    categories = ['Durability', 'Offense', 'Crowd Control', 'Mobility', 'Wave Control']
    
    # 1. Calculate Stat Gaps (Only used for Tie-Breaking or Reason Text now)
    stat_gaps = {}
    for cat in categories:
        my_val = my_stats.get(cat, 0)
        enemy_val = enemy_stats.get(cat, 0)
        gap = enemy_val - my_val
        stat_gaps[cat] = max(0, gap)

    # 2. Logic Split: Add vs Swap
    is_full_draft = len(my_team) >= 5
    
    # ---------------------------------------------------------
    # LOGIC A: INCOMPLETE DRAFT (ADD) - Simulation Based
    # ---------------------------------------------------------
    if not is_full_draft:
        # A1. Analyze Current State for Constraints
        current_mages = sum(1 for h in my_team if h in hero_data and "Mage" in hero_data[h]['roles'])
        current_mms = sum(1 for h in my_team if h in hero_data and "Marksman" in hero_data[h]['roles'])
        current_lanes = set()
        for h in my_team:
            if h in hero_data:
                for lane in hero_data[h]['lanes']:
                    if lane and lane.lower() != 'n/a':
                        current_lanes.add(lane)

        standard_lanes = ["Gold Lane", "EXP Lane", "Mid Lane", "Jungle", "Roaming"]
        missing_lanes = [l for l in standard_lanes if l not in current_lanes]

        # A2. Find Candidates using Simulation
        candidates = []
        
        for name in meta_scores.keys():
            if name in picked: continue
            if name not in hero_data: continue

            data = hero_data[name]
            candidate_roles = data['roles']
            candidate_lanes = data['lanes']

            # CONSTRAINTS CHECK
            valid = True
            
            # Role Constraints
            if current_mages == 0:
                if "Mage" not in candidate_roles: valid = False
            elif current_mages >= 1:
                if "Mage" in candidate_roles: valid = False

            if current_mms == 0:
                if "Marksman" not in candidate_roles: valid = False
            elif current_mms >= 1:
                if "Marksman" in candidate_roles: valid = False

            # Lane Constraints (Must fit a missing lane)
            if missing_lanes:
                if not any(l in candidate_lanes for l in missing_lanes):
                    valid = False

            if not valid:
                continue

            # SIMULATION: What if we add this hero?
            temp_team = list(my_team) + [name]
            temp_stats = calculate_weighted_stats(temp_team)
            projected_win_prob = evaluate_matchup_score(temp_stats, enemy_stats)

            # Score = Projected Win Rate
            # (We want the move that gives us the highest chance to win)
            candidates.append({
                "name": name,
                "score": projected_win_prob,
                "meta_score": meta_scores.get(name, 0)
            })

        if candidates:
            candidates.sort(key=lambda x: x['score'], reverse=True)
            best = candidates[0]
            
            # Generate Reason
            gap_cat = max(stat_gaps, key=stat_gaps.get)
            reason = f"Suggest {best['name']} (Meta: {best['meta_score']}). Maximizes win probability ({best['score']:.1%}) and fills required role/lane."
            
            return {
                "team": my_side_name,
                "type": "add",
                "heroes": [best['name']],
                "reason": reason
            }

    # ---------------------------------------------------------
    # LOGIC B: FULL DRAFT (SWAP) - Simulation
    # ---------------------------------------------------------
    else:
        best_swap_result = None
        highest_win_prob = -1

        # B1. Iterate over current team members
        for h_to_remove in my_team:
            # Determine constraints lost by removing this hero
            data_remove = hero_data[h_to_remove]
            remove_lanes = data_remove['lanes']
            remove_roles = data_remove['roles']
            
            strict_lanes_required = [l for l in remove_lanes if l and l.lower() != 'n/a' and l.lower() != 'nan']
            required_roles = []
            
            count_mage = sum(1 for h in my_team if h in hero_data and "Mage" in hero_data[h]['roles'])
            if "Mage" in remove_roles and count_mage == 1: required_roles.append("Mage")
                
            count_mm = sum(1 for h in my_team if h in hero_data and "Marksman" in hero_data[h]['roles'])
            if "Marksman" in remove_roles and count_mm == 1: required_roles.append("Marksman")

            # B2. Iterate Candidates
            for h_candidate in meta_scores.keys():
                if h_candidate in enemy_team: continue
                if h_candidate not in hero_data: continue
                if h_candidate in my_team: continue # Don't swap for same hero

                data_candidate = hero_data[h_candidate]
                candidate_lanes = data_candidate['lanes']
                candidate_roles = data_candidate['roles']

                # Constraints Check
                role_match = any(r in candidate_roles for r in required_roles) if required_roles else True
                lane_match = any(l in candidate_lanes for l in strict_lanes_required) if strict_lanes_required else True

                if not (role_match and lane_match):
                    continue

                # B3. Simulate Swap
                temp_team = [h for h in my_team if h != h_to_remove] + [h_candidate]
                
                # B4. Validate Composition
                if not validate_team_composition(temp_team):
                    continue

                # B5. Calculate Stats & Win Prob
                temp_stats = calculate_weighted_stats(temp_team)
                sim_win_prob = evaluate_matchup_score(temp_stats, enemy_stats)
                current_win_prob = evaluate_matchup_score(my_stats, enemy_stats)
                
                if sim_win_prob > highest_win_prob:
                    highest_win_prob = sim_win_prob
                    best_swap_result = {
                        "swap_out": h_to_remove,
                        "swap_in": h_candidate,
                        "new_prob": sim_win_prob,
                        "old_prob": current_win_prob
                    }

        # B6. Return Best Swap
        if best_swap_result:
            reason = (f"Swap {best_swap_result['swap_out']} for {best_swap_result['swap_in']}. "
                       f"Increases win probability from {best_swap_result['old_prob']:.1%} to {best_swap_result['new_prob']:.1%}.")
            return {
                "team": my_side_name,
                "type": "swap",
                "swap_out": best_swap_result['swap_out'],
                "swap_in": best_swap_result['swap_in'],
                "reason": reason
            }
        
        return {
            "team": my_side_name,
            "type": "swap",
            "swap_out": "None",
            "swap_in": "None",
            "reason": "Current composition is optimized based on Meta, Roles, Lanes, and Stats."
        }

    return None

# --- ROUTES ---

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/api/heroes', methods=['GET'])
def get_heroes():
    return jsonify(hero_data)

@app.route('/api/analyze', methods=['POST'])
def analyze():
    data = request.json
    blue_team = data.get('blue_team', [])
    red_team = data.get('red_team', [])

    # 1. Calculate WEIGHTED Stats
    blue_stats = calculate_weighted_stats(blue_team)
    red_stats = calculate_weighted_stats(red_team)

    # 2. Calculate Win Probability
    blue_power = sum(blue_stats.values())
    red_power = sum(red_stats.values())
    total_power = blue_power + red_power

    if total_power == 0:
        blue_prob = 50.0
        red_prob = 50.0
    else:
        blue_prob = (blue_power / total_power) * 100
        red_prob = 100.0 - blue_prob

    # 3. Prepare Stats Table Data
    stats_table_data = []
    categories = ['Durability', 'Offense', 'Crowd Control', 'Mobility', 'Wave Control']
    
    for cat in categories:
        b_val = blue_stats[cat]
        r_val = red_stats[cat]
        diff = abs(b_val - r_val)
        
        if b_val > r_val:
            leader = "blue"
        elif r_val > b_val:
            leader = "red"
        else:
            leader = "tie"

        stats_table_data.append({
            "stat": cat,
            "blue_val": round(b_val, 1),
            "red_val": round(r_val, 1),
            "diff": round(diff, 1),
            "leader": leader
        })

    # 4. Determine Advantages
    blue_adv = [row['stat'] for row in stats_table_data if row['leader'] == 'blue']
    red_adv = [row['stat'] for row in stats_table_data if row['leader'] == 'red']

    # 5. Get Smart Suggestion
    suggestion = None
    
    # Tolerance to avoid flip-flopping
    if abs(blue_prob - red_prob) > 0.5:
        if blue_prob < red_prob:
            suggestion = get_smart_suggestion(blue_team, red_team, blue_stats, red_stats, "Blue")
        elif red_prob < blue_prob:
            suggestion = get_smart_suggestion(red_team, blue_team, red_stats, blue_stats, "Red")

    return jsonify({
        "blue_prob": blue_prob,
        "red_prob": red_prob,
        "blue_adv": blue_adv,
        "red_adv": red_adv,
        "stats_table": stats_table_data,
        "suggestion": suggestion
    })

@app.route('/api/tournament-stats', methods=['GET'])
def get_tournament_stats():
    tournament_filter = request.args.get('tournament', 'All')
    map_filter = request.args.get('map', 'All')
    team_filter = request.args.get('team', 'All')
    side_filter = request.args.get('side', 'Blue')
    sort_by = request.args.get('sort', 'Total Picks')

    df = analyzer.get_hero_summary(
        side=side_filter,
        tournament_filter=tournament_filter,
        map_filter=map_filter,
        team_filter=team_filter
    )

    if not df.empty and sort_by in df.columns:
        df = df.sort_values(by=sort_by, ascending=False)

    return jsonify(df.to_dict(orient='records'))

@app.route('/api/filters', methods=['GET'])
def get_filters():
    current_tournament = request.args.get('tournament', 'All')

    tournaments = ["All"] + analyzer.get_unique_values('tournament')
    maps = ["All"] + analyzer.get_unique_values('map')

    if current_tournament and current_tournament != "All":
        teams = ["All"] + analyzer.get_unique_values('team', tournament_filter=current_tournament)
    else:
        teams = ["All"] + analyzer.get_unique_values('team')

    return jsonify({
        "tournaments": tournaments,
        "maps": maps,
        "teams": teams
    })

@app.route('/api/matches', methods=['GET'])
def get_matches():
    raw_matches = analyzer.matches
    matches_list = []

    for idx, match in enumerate(raw_matches):
        match_data = {
            "id": idx + 1,
            "tournament": match.get('tournament'),
            "map": match.get('map'),
            "blue_team": match.get('blue_team'),
            "red_team": match.get('red_team'),
            "blue_result": match.get('blue_result'),
            "red_result": match.get('red_result'),
            "blue_bans_p1": match.get('blue_bans', {}).get('p1', []),
            "blue_bans_p2": match.get('blue_bans', {}).get('p2', []),
            "blue_picks_p1": match.get('blue_picks', {}).get('p1', []),
            "blue_picks_p2": match.get('blue_picks', {}).get('p2', []),
            "red_bans_p1": match.get('red_bans', {}).get('p1', []),
            "red_bans_p2": match.get('red_bans', {}).get('p2', []),
            "red_picks_p1": match.get('red_picks', {}).get('p1', []),
            "red_picks_p2": match.get('red_picks', {}).get('p2', []),
        }
        matches_list.append(match_data)

    return jsonify(matches_list)

load_data()

if __name__ == '__main__':
    app.run(debug=True, port=8000)