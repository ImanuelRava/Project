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
    random.seed(name)
    r = lambda: random.randint(50, 200)
    return '#%02X%02X%02X' % (r(), r(), r())

def load_data():
    global hero_data, available_heroes_list
    if not os.path.exists(EXCEL_FILE):
        print(f"Warning: {EXCEL_FILE} not found. Hero stats will be missing.")
        return
    df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
    df.replace('N/A', pd.NA, inplace=True)
    stats_columns = ['Durability', 'Offense', 'Crowd Control', 'Mobility', 'Wave Control']

    for index, row in df.iterrows():
        name = row.get('Hero')
        if pd.isna(name): continue
        name = str(name).strip()
        roles = [str(row.get('Role 1')).strip()]
        role2 = row.get('Role 2')
        if pd.notna(role2) and str(role2).strip().lower() != 'nan': roles.append(str(role2).strip())
        lanes = [str(row.get('Lane 1')).strip()]
        lane2 = row.get('Lane 2')
        if pd.notna(lane2) and str(lane2).strip().lower() != 'nan': lanes.append(str(lane2).strip())
        stats = {}
        for stat in stats_columns: stats[stat] = float(row.get(stat)) if pd.notna(row.get(stat)) else 0.0
        hero_data[name] = {"roles": roles, "lanes": lanes, "stats": stats, "color": generate_color(name), "image": f"/static/hero_icon/{name}.png"}
        available_heroes_list.append(name)
    print(f"Loaded {len(hero_data)} heroes from {SHEET_NAME}.")

def calculate_weighted_stats(team_list):
    categories = ['Durability', 'Offense', 'Crowd Control', 'Mobility', 'Wave Control']
    weighted_stats = {cat: 0.0 for cat in categories}
    if not team_list: return weighted_stats
    for category in categories:
        hero_values = [hero_data[h]['stats'].get(category, 0) for h in team_list if h in hero_data]
        if not hero_values: continue
        max_val = max(hero_values)
        total_sum = sum(hero_values)
        count_others = len(hero_values) - 1
        avg_others = (total_sum - max_val) / count_others if count_others > 0 else 0
        weighted_stats[category] = (max_val * 0.4) + (avg_others * 0.6)
    return weighted_stats

def validate_team_composition(team_list):
    """Returns True if the team adheres to: 1 Mage, 1 Marksman, 5 Lanes."""
    roles_count = {"Mage": 0, "Marksman": 0}
    lanes_covered = set()
    required_lanes = ["Gold Lane", "EXP Lane", "Mid Lane", "Jungle", "Roaming"]
    for h in team_list:
        if h in hero_data:
            for role in hero_data[h]['roles']:
                if role in roles_count: roles_count[role] += 1
            for lane in hero_data[h]['lanes']:
                if lane and lane.lower() != 'n/a': lanes_covered.add(lane)
    return roles_count["Mage"] == 1 and roles_count["Marksman"] == 1 and all(l in lanes_covered for l in required_lanes)

def get_meta_popularity():
    popularity = {}
    if not analyzer.matches: return popularity
    for match in analyzer.matches:
        all_hero_lists = []
        def safe_extend(t, s):
            if s:
                p1, p2 = s.get('p1', []), s.get('p2', [])
                t.extend([h for h in p1 + p2 if h and str(h).strip() != ''])
        safe_extend(all_hero_lists, match.get('blue_picks'))
        safe_extend(all_hero_lists, match.get('red_picks'))
        safe_extend(all_hero_lists, match.get('blue_bans'))
        safe_extend(all_hero_lists, match.get('red_bans'))
        for hero in all_hero_lists: popularity[hero] = popularity.get(hero, 0) + 1
    return popularity

def evaluate_matchup_score(my_stats, enemy_stats):
    total_power = sum(my_stats.values()) + sum(enemy_stats.values())
    return sum(my_stats.values()) / total_power if total_power > 0 else 0.5

def get_smart_suggestion(my_team, enemy_team, my_stats, enemy_stats, my_side_name):
    picked = set(my_team + enemy_team)
    meta_scores = get_meta_popularity()
    categories = ['Durability', 'Offense', 'Crowd Control', 'Mobility', 'Wave Control']
    stat_gaps = {cat: max(0, enemy_stats.get(cat, 0) - my_stats.get(cat, 0)) for cat in categories}
    is_full_draft = len(my_team) >= 5
    
    if not is_full_draft:
        current_mages = sum(1 for h in my_team if h in hero_data and "Mage" in hero_data[h]['roles'])
        current_mms = sum(1 for h in my_team if h in hero_data and "Marksman" in hero_data[h]['roles'])
        current_lanes = set()
        for h in my_team:
            if h in hero_data:
                for lane in hero_data[h]['lanes']:
                    if lane and lane.lower() != 'n/a': current_lanes.add(lane)
        standard_lanes = ["Gold Lane", "EXP Lane", "Mid Lane", "Jungle", "Roaming"]
        missing_lanes = [l for l in standard_lanes if l not in current_lanes]
        candidates = []
        
        for name in meta_scores.keys():
            if name in picked or name not in hero_data: continue
            data = hero_data[name]
            candidate_roles = data['roles']
            candidate_lanes = data['lanes']
            valid = True
            
            if current_mages == 0:
                if "Mage" not in candidate_roles: valid = False
            elif current_mages >= 1:
                if "Mage" in candidate_roles: valid = False
            if current_mms == 0:
                if "Marksman" not in candidate_roles: valid = False
            elif current_mms >= 1:
                if "Marksman" in candidate_roles: valid = False
            if missing_lanes:
                if not any(l in candidate_lanes for l in missing_lanes): valid = False
            if not valid: continue
            
            temp_team = list(my_team) + [name]
            temp_stats = calculate_weighted_stats(temp_team)
            projected_win_prob = evaluate_matchup_score(temp_stats, enemy_stats)
            candidates.append({"name": name, "score": projected_win_prob, "meta_score": meta_scores.get(name, 0)})

        if candidates:
            candidates.sort(key=lambda x: x['score'], reverse=True)
            best = candidates[0]
            return {"team": my_side_name, "type": "add", "heroes": [best['name']], "reason": f"Suggest {best['name']} (Meta: {best['meta_score']}). Maximizes win probability ({best['score']:.1%})."}
    
    else:
        best_swap_result = None
        highest_win_prob = -1
        for h_to_remove in my_team:
            data_remove = hero_data[h_to_remove]
            remove_lanes = data_remove['lanes']
            remove_roles = data_remove['roles']
            strict_lanes_required = [l for l in remove_lanes if l and l.lower() != 'n/a' and l.lower() != 'nan']
            required_roles = []
            count_mage = sum(1 for h in my_team if h in hero_data and "Mage" in hero_data[h]['roles'])
            if "Mage" in remove_roles and count_mage == 1: required_roles.append("Mage")
            count_mm = sum(1 for h in my_team if h in hero_data and "Marksman" in hero_data[h]['roles'])
            if "Marksman" in remove_roles and count_mm == 1: required_roles.append("Marksman")
            
            for h_candidate in meta_scores.keys():
                if h_candidate in enemy_team or h_candidate not in hero_data or h_candidate in my_team: continue
                data_candidate = hero_data[h_candidate]
                candidate_lanes = data_candidate['lanes']
                candidate_roles = data_candidate['roles']
                role_match = any(r in candidate_roles for r in required_roles) if required_roles else True
                lane_match = any(l in candidate_lanes for l in strict_lanes_required) if strict_lanes_required else True
                if not (role_match and lane_match): continue
                
                temp_team = [h for h in my_team if h != h_to_remove] + [h_candidate]
                if not validate_team_composition(temp_team): continue
                
                temp_stats = calculate_weighted_stats(temp_team)
                sim_win_prob = evaluate_matchup_score(temp_stats, enemy_stats)
                current_win_prob = evaluate_matchup_score(my_stats, enemy_stats)
                
                if sim_win_prob > highest_win_prob:
                    highest_win_prob = sim_win_prob
                    best_swap_result = {"swap_out": h_to_remove, "swap_in": h_candidate, "new_prob": sim_win_prob, "old_prob": current_win_prob}
        
        if best_swap_result:
            return {"team": my_side_name, "type": "swap", "swap_out": best_swap_result['swap_out'], "swap_in": best_swap_result['swap_in'], "reason": f"Swap {best_swap_result['swap_out']} for {best_swap_result['swap_in']}. Prob: {best_swap_result['old_prob']:.1%} ➜ {best_swap_result['new_prob']:.1%}"}
        
        return {"team": my_side_name, "type": "swap", "swap_out": "None", "swap_in": "None", "reason": "Composition optimized."}

    return None

# --- ROUTES ---

@app.route('/')
def home(): return render_template('index.html')

@app.route('/api/heroes', methods=['GET'])
def get_heroes(): return jsonify(hero_data)

# Hero Relations (Synergy/Counters)
@app.route('/api/hero-relations/<hero_name>', methods=['GET'])
def get_hero_relations(hero_name):
    if not analyzer.is_loaded: return jsonify({"error": "Data not loaded"}), 503
    if hero_name not in hero_data: return jsonify({"error": "Hero not found"}), 404
    return jsonify({
        "hero": hero_name,
        "synergy": analyzer.get_synergy_for_hero(hero_name),
        "counters": analyzer.get_counters_for_hero(hero_name)
    })

@app.route('/api/analyze', methods=['POST'])
def analyze():
    data = request.json
    blue_team = data.get('blue_team', [])
    red_team = data.get('red_team', [])
    blue_stats = calculate_weighted_stats(blue_team)
    red_stats = calculate_weighted_stats(red_team)
    blue_power, red_power = sum(blue_stats.values()), sum(red_stats.values())
    total_power = blue_power + red_power
    blue_prob, red_prob = (blue_power / total_power * 100, 100 - (blue_power / total_power * 100)) if total_power > 0 else (50.0, 50.0)

    stats_table_data = []
    categories = ['Durability', 'Offense', 'Crowd Control', 'Mobility', 'Wave Control']
    for cat in categories:
        b_val, r_val = blue_stats[cat], red_stats[cat]
        diff = abs(b_val - r_val)
        leader = "blue" if b_val > r_val else "red" if r_val > b_val else "tie"
        stats_table_data.append({"stat": cat, "blue_val": round(b_val, 1), "red_val": round(r_val, 1), "diff": round(diff, 1), "leader": leader})
    
    blue_adv = [row['stat'] for row in stats_table_data if row['leader'] == 'blue']
    red_adv = [row['stat'] for row in stats_table_data if row['leader'] == 'red']

    # Smart Suggestion
    suggestion = None
    if abs(blue_prob - red_prob) > 0.5:
        if blue_prob < red_prob: suggestion = get_smart_suggestion(blue_team, red_team, blue_stats, red_stats, "Blue")
        elif red_prob < blue_prob: suggestion = get_smart_suggestion(red_team, blue_team, red_stats, blue_stats, "Red")

    return jsonify({
        "blue_prob": blue_prob, "red_prob": red_prob, 
        "blue_adv": blue_adv, "red_adv": red_adv, 
        "stats_table": stats_table_data, "suggestion": suggestion
    })

# --- NEW ROUTES ---

@app.route('/api/suggest-bans', methods=['POST'])
def suggest_bans():
    """
    Suggests bans based on:
    1. Heroes that counter the user's team (Threats).
    2. High Meta heroes (Popularity).
    """
    data = request.json
    my_team = data.get('my_team', []) # List of hero names
    enemy_team = data.get('enemy_team', [])
    
    # 1. Get threats (Heroes that beat my team)
    threats = analyzer.get_counters_for_team(my_team)
    
    # 2. Get Meta Popularity
    meta_scores = get_meta_popularity()
    
    ban_scores = {}
    
    # Score Logic
    for hero, score in meta_scores.items():
        if hero in my_team or hero in enemy_team: continue
        
        base_score = score # Popularity score
        
        # Check if this hero is a threat to my team
        threat_info = next((t for t in threats if t['hero'] == hero), None)
        threat_bonus = 0
        reason = "Meta Priority"
        
        if threat_info:
            # If they beat my team, they are HIGH priority bans
            threat_bonus = (threat_info['win_rate'] - 50) * 5 # Scale by how much they beat us
            reason = "Counter Threat"
        
        ban_scores[hero] = base_score + threat_bonus

    # Sort by score descending
    sorted_bans = sorted(ban_scores.items(), key=lambda x: x[1], reverse=True)
    
    top_5 = []
    for hero, score in sorted_bans[:5]:
        # Determine reason for display
        t_info = next((t for t in threats if t['hero'] == hero), None)
        display_reason = "Meta Priority"
        if t_info and t_info['win_rate'] > 55:
            display_reason = f"Counter Threat ({t_info['win_rate']}% WR vs You)"
            
        top_5.append({'name': hero, 'reason': display_reason, 'score': score})
        
    return jsonify({"suggestions": top_5})

@app.route('/api/team-signatures/<team_name>', methods=['GET'])
def get_team_signatures(team_name):
    if not analyzer.is_loaded: return jsonify({"error": "Data not loaded"}), 503
    return jsonify(analyzer.get_team_signatures(team_name))

@app.route('/api/meta-summary', methods=['GET'])
def get_meta_summary():
    """
    Returns global meta stats: Top Picks, Top Bans, Highest Win Rates.
    """
    if not analyzer.is_loaded: return jsonify({"error": "Data not loaded"}), 503
    
    all_stats = analyzer.get_global_meta_stats()
    
    # Sort by different metrics
    top_picks = sorted(all_stats, key=lambda x: x['picks'], reverse=True)[:5]
    top_bans = sorted(all_stats, key=lambda x: x['bans'], reverse=True)[:5]
    top_winrate = [h for h in sorted(all_stats, key=lambda x: x['win_rate'], reverse=True) if h['picks'] >= 5][:5] # Min 5 games to qualify
    
    return jsonify({
        "top_picks": top_picks,
        "top_bans": top_bans,
        "highest_winrate": top_winrate
    })

@app.route('/api/tournament-stats', methods=['GET'])
def get_tournament_stats():
    df = analyzer.get_hero_summary(side=request.args.get('side', 'Blue'), tournament_filter=request.args.get('tournament', 'All'), map_filter=request.args.get('map', 'All'), team_filter=request.args.get('team', 'All'))
    sort_by = request.args.get('sort', 'Total Picks')
    if not df.empty and sort_by in df.columns: df = df.sort_values(by=sort_by, ascending=False)
    return jsonify(df.to_dict(orient='records'))

@app.route('/api/filters', methods=['GET'])
def get_filters():
    tourn_filter = request.args.get('tournament', 'All')
    tournaments = ["All"] + analyzer.get_unique_values('tournament')
    maps = ["All"] + analyzer.get_unique_values('map')
    teams = ["All"] + analyzer.get_unique_values('team', tournament_filter=tourn_filter) if tourn_filter != "All" else ["All"] + analyzer.get_unique_values('team')
    return jsonify({"tournaments": tournaments, "maps": maps, "teams": teams})

@app.route('/api/matches', methods=['GET'])
def get_matches():
    raw_matches = analyzer.matches
    matches_list = []
    f_team, f_enemy, f_tourn = request.args.get('team'), request.args.get('enemy'), request.args.get('tournament')
    
    for idx, match in enumerate(raw_matches):
        if f_team and match['blue_team'] != f_team and match['red_team'] != f_team: continue
        if f_enemy:
            is_b = match['blue_team'] == f_team
            is_r = match['red_team'] == f_team
            if is_b and match['red_team'] != f_enemy: continue
            if is_r and match['blue_team'] != f_enemy: continue
        if f_tourn and match['tournament'] != f_tourn: continue
        
        matches_list.append({
            "id": idx + 1, "tournament": match.get('tournament'), "map": match.get('map'),
            "blue_team": match.get('blue_team'), "red_team": match.get('red_team'),
            "blue_result": match.get('blue_result'), "red_result": match.get('red_result'),
            "blue_bans_p1": match.get('blue_bans', {}).get('p1', []), "blue_bans_p2": match.get('blue_bans', {}).get('p2', []),
            "blue_picks_p1": match.get('blue_picks', {}).get('p1', []), "blue_picks_p2": match.get('blue_picks', {}).get('p2', []),
            "red_bans_p1": match.get('red_bans', {}).get('p1', []), "red_bans_p2": match.get('red_bans', {}).get('p2', []),
            "red_picks_p1": match.get('red_picks', {}).get('p1', []), "red_picks_p2": match.get('red_picks', {}).get('p2', []),
        })
    return jsonify(matches_list)

load_data()
if __name__ == '__main__':
    app.run(debug=True, port=8000)