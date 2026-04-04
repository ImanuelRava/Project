import pandas as pd
import random
import os
import logging
from flask import Flask, jsonify, request, render_template
from flask_cors import CORS
from data_analyzer import DataAnalyzer

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE = os.path.join(BASE_DIR, 'Analyst.xlsx')
SHEET_NAME = 'Hero Data'
STAT_CATEGORIES = ['Durability', 'Offense', 'Crowd Control', 'Mobility', 'Wave Control']
REQUIRED_LANES = ["Gold Lane", "EXP Lane", "Mid Lane", "Jungle", "Roaming"]

app = Flask(__name__)
CORS(app)

hero_data = {}
available_heroes_list = []
analyzer = None

def generate_color(name):
    random.seed(name)
    r = lambda: random.randint(50, 200)
    return '#%02X%02X%02X' % (r(), r(), r())

def load_data():
    global hero_data, available_heroes_list, analyzer
    analyzer = DataAnalyzer()

    if not os.path.exists(EXCEL_FILE):
        logger.warning(f"{EXCEL_FILE} not found. Hero stats will be missing.")
        return
    
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
        df.replace('N/A', pd.NA, inplace=True)
        
        for index, row in df.iterrows():
            name = row.get('Hero')
            if pd.isna(name): continue
            name = str(name).strip()
            roles = _extract_roles(row)
            lanes = _extract_lanes(row)
            stats = _extract_stats(row)
            
            hero_data[name] = {
                "roles": roles, "lanes": lanes, "stats": stats,
                "color": generate_color(name), "image": f"/static/hero_icon/{name}.png"
            }
            available_heroes_list.append(name)
        
        logger.info(f"Loaded {len(hero_data)} heroes from {SHEET_NAME}.")
        analyzer.set_hero_data(hero_data)
        
        mismatches = analyzer.get_hero_name_mismatches()
        if mismatches:
            logger.warning(f"Found {len(mismatches)} hero name mismatches:")
            for m in mismatches[:10]: logger.warning(f"  - {m}")
    except Exception as e:
        logger.error(f"Error loading hero data: {e}")

def _extract_roles(row):
    roles = [str(row.get('Role 1')).strip()]
    role2 = row.get('Role 2')
    if pd.notna(role2) and str(role2).strip().lower() != 'nan': roles.append(str(role2).strip())
    return roles

def _extract_lanes(row):
    lanes = [str(row.get('Lane 1')).strip()]
    lane2 = row.get('Lane 2')
    if pd.notna(lane2) and str(lane2).strip().lower() != 'nan': lanes.append(str(lane2).strip())
    return lanes

def _extract_stats(row):
    return {stat: float(row.get(stat)) if pd.notna(row.get(stat)) else 0.0 for stat in STAT_CATEGORIES}

def calculate_weighted_stats(team_list):
    weighted_stats = {cat: 0.0 for cat in STAT_CATEGORIES}
    if not team_list: return weighted_stats
    for category in STAT_CATEGORIES:
        hero_values = [hero_data[h]['stats'].get(category, 0) for h in team_list if h in hero_data]
        if not hero_values: continue
        max_val = max(hero_values)
        total_sum = sum(hero_values)
        count_others = len(hero_values) - 1
        avg_others = (total_sum - max_val) / count_others if count_others > 0 else 0
        weighted_stats[category] = (max_val * 0.4) + (avg_others * 0.6)
    return weighted_stats

def validate_team_composition(team_list):
    roles_count = {"Mage": 0, "Marksman": 0}
    lanes_covered = set()
    for h in team_list:
        if h not in hero_data: continue
        for role in hero_data[h]['roles']:
            if role in roles_count: roles_count[role] += 1
        for lane in hero_data[h]['lanes']:
            if lane and lane.lower() != 'n/a': lanes_covered.add(lane)
    return roles_count["Mage"] == 1 and roles_count["Marksman"] == 1 and all(l in lanes_covered for l in REQUIRED_LANES)

def get_meta_popularity():
    if not analyzer or not analyzer.matches: return {}
    popularity = {}
    for match in analyzer.matches:
        all_heroes = _extract_all_heroes_from_match(match)
        for hero in all_heroes: popularity[hero] = popularity.get(hero, 0) + 1
    return popularity

def _extract_all_heroes_from_match(match):
    heroes = []
    for side in ['blue', 'red']:
        for phase_type in ['picks', 'bans']:
            phase_data = match.get(f'{side}_{phase_type}', {})
            if phase_data: heroes.extend([h for h in phase_data.get('p1', []) + phase_data.get('p2', []) if h and str(h).strip()])
    return heroes

def evaluate_matchup_score(my_stats, enemy_stats):
    total_power = sum(my_stats.values()) + sum(enemy_stats.values())
    return sum(my_stats.values()) / total_power if total_power > 0 else 0.5

def get_smart_suggestion(my_team, enemy_team, my_stats, enemy_stats, my_side_name):
    picked = set(my_team + enemy_team)
    meta_scores = get_meta_popularity()
    if len(my_team) < 5: return _get_add_suggestion(my_team, enemy_team, my_stats, enemy_stats, my_side_name, picked, meta_scores)
    else: return _get_swap_suggestion(my_team, enemy_team, my_stats, enemy_stats, my_side_name, picked, meta_scores)

def _get_add_suggestion(my_team, enemy_team, my_stats, enemy_stats, my_side_name, picked, meta_scores):
    current_mages = _count_role(my_team, "Mage")
    current_mms = _count_role(my_team, "Marksman")
    current_lanes = _get_covered_lanes(my_team)
    missing_lanes = [l for l in REQUIRED_LANES if l not in current_lanes]
    candidates = []
    for name in meta_scores.keys():
        if name in picked or name not in hero_data: continue
        if not _is_valid_addition(name, current_mages, current_mms, missing_lanes): continue
        temp_team = list(my_team) + [name]
        temp_stats = calculate_weighted_stats(temp_team)
        projected_win_prob = evaluate_matchup_score(temp_stats, enemy_stats)
        candidates.append({"name": name, "score": projected_win_prob, "meta_score": meta_scores.get(name, 0)})
    if candidates:
        candidates.sort(key=lambda x: x['score'], reverse=True)
        best = candidates[0]
        return {"team": my_side_name, "type": "add", "heroes": [best['name']], "reason": f"Suggest {best['name']} (Meta: {best['meta_score']}). Maximizes win probability ({best['score']:.1%})."}
    return None

def _get_swap_suggestion(my_team, enemy_team, my_stats, enemy_stats, my_side_name, picked, meta_scores):
    best_swap_result = None
    highest_win_prob = -1
    current_win_prob = evaluate_matchup_score(my_stats, enemy_stats)
    for h_to_remove in my_team:
        swap_candidates = _find_swap_candidates(h_to_remove, my_team, enemy_team, meta_scores)
        for h_candidate in swap_candidates:
            temp_team = [h for h in my_team if h != h_to_remove] + [h_candidate]
            if not validate_team_composition(temp_team): continue
            temp_stats = calculate_weighted_stats(temp_team)
            sim_win_prob = evaluate_matchup_score(temp_stats, enemy_stats)
            if sim_win_prob > highest_win_prob:
                highest_win_prob = sim_win_prob
                best_swap_result = {"swap_out": h_to_remove, "swap_in": h_candidate, "new_prob": sim_win_prob, "old_prob": current_win_prob}
    if best_swap_result:
        return {"team": my_side_name, "type": "swap", "swap_out": best_swap_result['swap_out'], "swap_in": best_swap_result['swap_in'], "reason": f"Swap {best_swap_result['swap_out']} for {best_swap_result['swap_in']}. Prob: {best_swap_result['old_prob']:.1%} ➜ {best_swap_result['new_prob']:.1%}"}
    return {"team": my_side_name, "type": "swap", "swap_out": "None", "swap_in": "None", "reason": "Composition optimized."}

def _count_role(team, role):
    return sum(1 for h in team if h in hero_data and role in hero_data[h]['roles'])

def _get_covered_lanes(team):
    lanes = set()
    for h in team:
        if h in hero_data:
            for lane in hero_data[h]['lanes']:
                if lane and lane.lower() != 'n/a': lanes.add(lane)
    return lanes

def _is_valid_addition(hero_name, current_mages, current_mms, missing_lanes):
    data = hero_data[hero_name]
    candidate_roles, candidate_lanes = data['roles'], data['lanes']
    if current_mages == 0 and "Mage" not in candidate_roles: return False
    if current_mages >= 1 and "Mage" in candidate_roles: return False
    if current_mms == 0 and "Marksman" not in candidate_roles: return False
    if current_mms >= 1 and "Marksman" in candidate_roles: return False
    if missing_lanes and not any(l in candidate_lanes for l in missing_lanes): return False
    return True

def _find_swap_candidates(h_to_remove, my_team, enemy_team, meta_scores):
    data_remove = hero_data[h_to_remove]
    remove_lanes = [l for l in data_remove['lanes'] if l and l.lower() not in ('n/a', 'nan')]
    remove_roles = data_remove['roles']
    required_roles = []
    if "Mage" in remove_roles and _count_role(my_team, "Mage") == 1: required_roles.append("Mage")
    if "Marksman" in remove_roles and _count_role(my_team, "Marksman") == 1: required_roles.append("Marksman")
    candidates = []
    for h_candidate in meta_scores.keys():
        if h_candidate in enemy_team or h_candidate not in hero_data or h_candidate in my_team: continue
        data_candidate = hero_data[h_candidate]
        role_match = any(r in data_candidate['roles'] for r in required_roles) if required_roles else True
        lane_match = any(l in data_candidate['lanes'] for l in remove_lanes) if remove_lanes else True
        if role_match and lane_match: candidates.append(h_candidate)
    return candidates

def _calculate_win_probabilities(blue_stats, red_stats):
    blue_power = sum(blue_stats.values())
    red_power = sum(red_stats.values())
    total_power = blue_power + red_power
    if total_power > 0: return blue_power / total_power * 100, 100 - (blue_power / total_power * 100)
    return 50.0, 50.0

def _build_stats_table(blue_stats, red_stats):
    stats_table_data = []
    for cat in STAT_CATEGORIES:
        b_val, r_val = blue_stats[cat], red_stats[cat]
        diff = abs(b_val - r_val)
        leader = "blue" if b_val > r_val else "red" if r_val > b_val else "tie"
        stats_table_data.append({"stat": cat, "blue_val": round(b_val, 1), "red_val": round(r_val, 1), "diff": round(diff, 1), "leader": leader})
    return stats_table_data

def _get_suggestion_if_needed(blue_team, red_team, blue_stats, red_stats, blue_prob, red_prob):
    if abs(blue_prob - red_prob) <= 0.5: return None
    if blue_prob < red_prob: return get_smart_suggestion(blue_team, red_team, blue_stats, red_stats, "Blue")
    else: return get_smart_suggestion(red_team, blue_team, red_stats, blue_stats, "Red")

def _calculate_ban_scores(my_team, enemy_team, meta_scores, threats):
    ban_scores = {}
    threat_lookup = {t['hero']: t for t in threats}
    for hero, score in meta_scores.items():
        if hero in my_team or hero in enemy_team: continue
        base_score = score
        threat_bonus = (threat_lookup[hero]['win_rate'] - 50) * 5 if hero in threat_lookup else 0
        ban_scores[hero] = base_score + threat_bonus
    return ban_scores

def _get_ban_reason(hero, threats):
    threat_info = next((t for t in threats if t['hero'] == hero), None)
    if threat_info and threat_info['win_rate'] > 55: return f"Counter Threat ({threat_info['win_rate']}% WR vs You)"
    return "Meta Priority"

def _passes_match_filters(match, f_team, f_enemy, f_tourn):
    if f_team and match['blue_team'] != f_team and match['red_team'] != f_team: return False
    if f_enemy:
        is_blue = match['blue_team'] == f_team
        is_red = match['red_team'] == f_team
        if is_blue and match['red_team'] != f_enemy: return False
        if is_red and match['blue_team'] != f_enemy: return False
    if f_tourn and match['tournament'] != f_tourn: return False
    return True

def _format_match_for_response(idx, match):
    return {
        "id": idx + 1, "tournament": match.get('tournament'), "map": match.get('map'),
        "blue_team": match.get('blue_team'), "red_team": match.get('red_team'),
        "blue_result": match.get('blue_result'), "red_result": match.get('red_result'),
        "blue_bans_p1": match.get('blue_bans', {}).get('p1', []), "blue_bans_p2": match.get('blue_bans', {}).get('p2', []),
        "blue_picks_p1": match.get('blue_picks', {}).get('p1', []), "blue_picks_p2": match.get('blue_picks', {}).get('p2', []),
        "red_bans_p1": match.get('red_bans', {}).get('p1', []), "red_bans_p2": match.get('red_bans', {}).get('p2', []),
        "red_picks_p1": match.get('red_picks', {}).get('p1', []), "red_picks_p2": match.get('red_picks', {}).get('p2', []),
    }

# --- ROUTES ---

@app.route('/')
def home(): return render_template('index.html')

@app.route('/api/heroes', methods=['GET'])
def get_heroes(): return jsonify(hero_data)

@app.route('/api/status', methods=['GET'])
def get_status():
    if not analyzer: return jsonify({"status": "error", "message": "Analyzer not initialized"})
    status = analyzer.get_load_status()
    status["flask_hero_data_count"] = len(hero_data)
    status["excel_file_exists"] = os.path.exists(EXCEL_FILE)
    return jsonify({"status": "ok" if status["is_loaded"] and status["hero_data_loaded"] else "warning", **status})

@app.route('/api/debug/heroes', methods=['GET'])
def debug_heroes():
    return jsonify({
        "hero_data_keys": sorted(list(hero_data.keys())),
        "match_heroes": analyzer._get_available_heroes_from_matches() if analyzer else [],
        "mismatches": analyzer.get_hero_name_mismatches() if analyzer else []
    })

@app.route('/api/debug/tournament-stats', methods=['GET'])
def debug_tournament_stats():
    """Debug endpoint to see what's happening with tournament stats."""
    try:
        if not analyzer:
            return jsonify({"error": "Analyzer not initialized"})
        
        sample_matches = []
        for i, match in enumerate(analyzer.matches[:3]):
            sample_matches.append({
                "index": i,
                "tournament": match.get('tournament'),
                "map": match.get('map'),
                "blue_team": match.get('blue_team'),
                "red_team": match.get('red_team'),
                "blue_picks_count": len([h for h in match.get('blue_picks', {}).get('p1', []) + match.get('blue_picks', {}).get('p2', []) if h]),
                "red_picks_count": len([h for h in match.get('red_picks', {}).get('p1', []) + match.get('red_picks', {}).get('p2', []) if h])
            })
        
        tournaments = analyzer.get_unique_values('tournament')
        maps = analyzer.get_unique_values('map')
        teams = analyzer.get_unique_values('team')
        
        test_df = analyzer.get_hero_summary(side='Blue', tournament_filter='All', map_filter='All', team_filter='All')
        
        return jsonify({
            "total_matches": len(analyzer.matches),
            "sample_matches": sample_matches,
            "available_tournaments": tournaments,
            "available_maps": maps,
            "available_teams": teams,
            "test_query_result_count": len(test_df),
            "test_query_columns": list(test_df.columns) if not test_df.empty else [],
            "is_loaded": analyzer.is_loaded,
            "load_errors": analyzer._load_errors
        })
    except Exception as e:
        return jsonify({"error": str(e)})

@app.route('/api/debug/analyze-test', methods=['GET'])
def debug_analyze_test():
    if not analyzer: return jsonify({"error": "Analyzer not initialized"}), 503
    test_heroes = list(hero_data.keys())[:5]
    if len(test_heroes) < 5: return jsonify({"error": f"Only {len(test_heroes)} heroes available"}), 400
    result = analyzer.analyze_comp_archetype(test_heroes, test_heroes[:5])
    return jsonify({"test_heroes_blue": test_heroes, "test_heroes_red": test_heroes[:5], "analysis_result": result, "hero_stats_sample": {h: hero_data[h]['stats'] for h in test_heroes if h in hero_data}})

@app.route('/api/hero-relations/<hero_name>', methods=['GET'])
def get_hero_relations(hero_name):
    if not analyzer or not analyzer.is_loaded: return jsonify({"error": "Data not loaded"}), 503
    if hero_name not in hero_data: return jsonify({"error": "Hero not found"}), 404
    return jsonify({"hero": hero_name, "synergy": analyzer.get_synergy_for_hero(hero_name), "counters": analyzer.get_counters_for_hero(hero_name)})

@app.route('/api/analyze', methods=['POST'])
def analyze():
    data = request.json
    blue_team = [h for h in data.get('blue_team', []) if h and str(h).strip()]
    red_team = [h for h in data.get('red_team', []) if h and str(h).strip()]
    invalid_blue = [h for h in blue_team if h not in hero_data]
    invalid_red = [h for h in red_team if h not in hero_data]
    blue_stats = calculate_weighted_stats(blue_team)
    red_stats = calculate_weighted_stats(red_team)
    blue_prob, red_prob = _calculate_win_probabilities(blue_stats, red_stats)
    stats_table_data = _build_stats_table(blue_stats, red_stats)
    blue_adv = [row['stat'] for row in stats_table_data if row['leader'] == 'blue']
    red_adv = [row['stat'] for row in stats_table_data if row['leader'] == 'red']
    suggestion = _get_suggestion_if_needed(blue_team, red_team, blue_stats, red_stats, blue_prob, red_prob)
    warnings = []
    if invalid_blue: warnings.append(f"Unknown Blue heroes: {invalid_blue}")
    if invalid_red: warnings.append(f"Unknown Red heroes: {invalid_red}")
    if len(blue_team) == 0 and len(red_team) == 0: warnings.append("No heroes selected for either team")
    return jsonify({"blue_prob": blue_prob, "red_prob": red_prob, "blue_adv": blue_adv, "red_adv": red_adv, "stats_table": stats_table_data, "suggestion": suggestion, "warnings": warnings, "blue_hero_count": len(blue_team), "red_hero_count": len(red_team)})

@app.route('/api/suggest-bans', methods=['POST'])
def suggest_bans():
    data = request.json
    my_team = [h for h in data.get('my_team', []) if h and str(h).strip()]
    enemy_team = [h for h in data.get('enemy_team', []) if h and str(h).strip()]
    if not analyzer or not analyzer.is_loaded: return jsonify({"suggestions": [], "error": "Match data not loaded"}), 503
    threats = analyzer.get_counters_for_team(my_team)
    meta_scores = get_meta_popularity()
    if not meta_scores: return jsonify({"suggestions": [], "error": "No meta data available"}), 503
    ban_scores = _calculate_ban_scores(my_team, enemy_team, meta_scores, threats)
    sorted_bans = sorted(ban_scores.items(), key=lambda x: x[1], reverse=True)
    top_5 = [{'name': hero, 'reason': _get_ban_reason(hero, threats), 'score': score} for hero, score in sorted_bans[:5]]
    return jsonify({"suggestions": top_5})

@app.route('/api/team-signatures/<team_name>', methods=['GET'])
def get_team_signatures(team_name):
    if not analyzer or not analyzer.is_loaded: return jsonify({"error": "Data not loaded"}), 503
    return jsonify(analyzer.get_team_signatures(team_name))

@app.route('/api/meta-summary', methods=['GET'])
def get_meta_summary():
    if not analyzer or not analyzer.is_loaded: return jsonify({"error": "Data not loaded"}), 503
    all_stats = analyzer.get_global_meta_stats()
    top_picks = sorted(all_stats, key=lambda x: x['picks'], reverse=True)[:5]
    top_bans = sorted(all_stats, key=lambda x: x['bans'], reverse=True)[:5]
    top_winrate = [h for h in sorted(all_stats, key=lambda x: x['win_rate'], reverse=True) if h['picks'] >= 5][:5]
    return jsonify({"top_picks": top_picks, "top_bans": top_bans, "highest_winrate": top_winrate})

@app.route('/api/tournament-stats', methods=['GET'])
def get_tournament_stats():
    """Get tournament hero statistics."""
    if not analyzer or not analyzer.is_loaded:
        return jsonify([]), 503
    
    try:
        df = analyzer.get_hero_summary(
            side=request.args.get('side', 'Blue'),
            tournament_filter=request.args.get('tournament', 'All'),
            map_filter=request.args.get('map', 'All'),
            team_filter=request.args.get('team', 'All')
        )
        
        sort_by = request.args.get('sort', 'Total Picks')
        if not df.empty and sort_by in df.columns:
            df = df.sort_values(by=sort_by, ascending=False)
        
        return jsonify(df.to_dict(orient='records'))
    except Exception as e:
        logger.error(f"Tournament stats error: {e}", exc_info=True)
        return jsonify([])

@app.route('/api/filters', methods=['GET'])
def get_filters():
    if not analyzer: return jsonify({"tournaments": [], "maps": [], "teams": []})
    tourn_filter = request.args.get('tournament', 'All')
    tournaments = ["All"] + analyzer.get_unique_values('tournament')
    maps = ["All"] + analyzer.get_unique_values('map')
    teams = ["All"] + analyzer.get_unique_values('team', tournament_filter=tourn_filter) if tourn_filter != "All" else ["All"] + analyzer.get_unique_values('team')
    return jsonify({"tournaments": tournaments, "maps": maps, "teams": teams})

@app.route('/api/matches', methods=['GET'])
def get_matches():
    if not analyzer: return jsonify([])
    matches_list = []
    f_team, f_enemy, f_tourn = request.args.get('team'), request.args.get('enemy'), request.args.get('tournament')
    for idx, match in enumerate(analyzer.matches):
        if _passes_match_filters(match, f_team, f_enemy, f_tourn): matches_list.append(_format_match_for_response(idx, match))
    return jsonify(matches_list)

@app.route('/api/meta-archetypes', methods=['GET'])
def get_meta_archetypes():
    if not analyzer or not analyzer.is_loaded: return jsonify({"error": "Match data not loaded"}), 503
    if not analyzer.hero_data_dict: return jsonify({"error": "Hero stats not loaded", "archetypes": []}), 503
    return jsonify(analyzer.get_archetype_stats())

@app.route('/api/analyze-comp', methods=['POST'])
def analyze_comp_archetype():
    if not analyzer or not analyzer.is_loaded: return jsonify({"error": "Match data not loaded", "details": analyzer._load_errors}), 503
    if not analyzer.hero_data_dict: return jsonify({"error": "Hero stats not loaded"}), 503
    data = request.json
    blue_team = [h for h in data.get('blue_team', []) if h and str(h).strip()]
    red_team = [h for h in data.get('red_team', []) if h and str(h).strip()]
    invalid_blue = [h for h in blue_team if h not in hero_data]
    invalid_red = [h for h in red_team if h not in hero_data]
    result = analyzer.analyze_comp_archetype(blue_team, red_team)
    result['warnings'] = []
    if invalid_blue: result['warnings'].append(f"Unknown heroes on Blue: {invalid_blue}")
    if invalid_red: result['warnings'].append(f"Unknown heroes on Red: {invalid_red}")
    if len(blue_team) < 5: result['warnings'].append(f"Blue team has {len(blue_team)}/5 heroes - stats are extrapolated")
    if len(red_team) < 5: result['warnings'].append(f"Red team has {len(red_team)}/5 heroes - stats are extrapolated")
    return jsonify(result)

load_data()

if __name__ == '__main__':
    logger.info("Starting Flask application...")
    logger.info(f"Analyzer loaded: {analyzer.is_loaded if analyzer else False}")
    logger.info(f"Hero data loaded: {len(hero_data)} heroes")
    app.run(debug=True, port=8000)