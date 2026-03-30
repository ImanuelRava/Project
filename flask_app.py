import pandas as pd
import math
import random
from flask import Flask, jsonify, request, render_template
from flask_cors import CORS
from data_analyzer import DataAnalyzer

# --- CONFIGURATION ---
EXCEL_FILE = 'Analyst.xlsx'
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

    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
    except FileNotFoundError:
        print(f"Error: Could not find {EXCEL_FILE}")
        return

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
            "color": generate_color(name), # Keep color as backup
            "image": image_path            # Add image path
        }
        available_heroes_list.append(name)

    print(f"Loaded {len(hero_data)} heroes from {SHEET_NAME}.")

def calculate_team_stats(team_list):
    """Calculates aggregate stats for a list of heroes."""
    total_stats = {
        'Durability': 0, 'Offense': 0, 'Crowd Control': 0,
        'Mobility': 0, 'Lane Control': 0
    }

    count = 0
    for hero_name in team_list:
        if hero_name in hero_data:
            h_stats = hero_data[hero_name]['stats']
            for key in total_stats:
                total_stats[key] += h_stats[key]
            count += 1

    avg_stats = {k: (v / count) if count > 0 else 0 for k, v in total_stats.items()}
    return avg_stats

def calculate_weighted_stats(team_list):
    categories = ['Durability', 'Offense', 'Crowd Control', 'Mobility', 'Wave Control']
    weighted_stats = {cat: 0.0 for cat in categories}

    if not team_list:
        return weighted_stats

    for category in categories:
        # 1. Collect the values for this specific stat from all selected heroes
        hero_values = []
        for hero_name in team_list:
            if hero_name in hero_data:
                val = hero_data[hero_name]['stats'].get(category, 0)
                hero_values.append(val)

        if not hero_values:
            weighted_stats[category] = 0
            continue

        # 2. Find the Highest Value (The "Carry" for this stat)
        max_val = max(hero_values)

        # 3. Calculate Average of the OTHER heroes
        total_sum = sum(hero_values)
        count_others = len(hero_values) - 1

        if count_others > 0:
            sum_others = total_sum - max_val
            avg_others = sum_others / count_others
        else:
            avg_others = 0

        # 4. Apply Formula: Max * 0.7 + Avg Others * 0.3
        weighted_score = (max_val * 0.7) + (avg_others * 0.3)
        weighted_stats[category] = weighted_score

    return weighted_stats

def generate_radar_chart(blue_stats, red_stats):
    """Generates an SVG string for the Radar Chart with non-overlapping labels."""
    width, height = 300, 300
    center = 150
    radius = 100
    labels = ['Durability', 'Offense', 'Crowd Control', 'Mobility', 'Wave Control']
    num_axes = len(labels)
    angle_step = (2 * math.pi) / num_axes

    # Helper to get coordinates for data points
    def get_coords(value, index):
        angle = index * angle_step - (math.pi / 2)
        val_norm = min(value / 10.0, 1.0)
        x = center + (radius * val_norm) * math.cos(angle)
        y = center + (radius * val_norm) * math.sin(angle)
        return x, y

    # Helper to get coordinates for LABELS (pushed further out)
    def get_label_coords(index):
        angle = index * angle_step - (math.pi / 2)
        # Push labels 25% further than the chart radius
        r = radius * 1.25
        x = center + r * math.cos(angle)
        y = center + r * math.sin(angle)
        return x, y

    # Generate Points
    blue_points = []
    red_points = []
    grid_points = []
    label_elements = []

    for i in range(num_axes):
        # 1. Data Points
        bx, by = get_coords(blue_stats[labels[i]], i)
        blue_points.append(f"{bx},{by}")

        rx, ry = get_coords(red_stats[labels[i]], i)
        red_points.append(f"{rx},{ry}")

        # 2. Grid Points
        gx, gy = get_coords(10, i)
        grid_points.append(f"{gx},{gy}")

        # 3. Labels with Alignment
        lx, ly = get_label_coords(i)

        # Determine text alignment based on X position to prevent overlap
        text_anchor = "middle"
        if lx < center:
            text_anchor = "end"  # Align right if on the left
        elif lx > center:
            text_anchor = "start" # Align left if on the right

        # Create text element with a background stroke (halo) for readability
        label_elements.append(
            f'<text x="{lx}" y="{ly}" '
            f'fill="#cbd5e1" font-size="11" font-weight="600" '
            f'text-anchor="{text_anchor}" dominant-baseline="middle" '
            f'stroke="#0f172a" stroke-width="3" paint-order="stroke fill">'
            f'{labels[i]}</text>'
        )

    # Build SVG
    svg = f"""
    <svg width="100%" height="100%" viewBox="0 0 {width} {height}" xmlns="http://www.w3.org/2000/svg">
        <!-- Background Grid -->
        <polygon points="{' '.join(grid_points)}" fill="#1e293b" stroke="#334155" stroke-width="1"/>

        <!-- Grid Circles -->
        <circle cx="{center}" cy="{center}" r="{radius*0.66}" fill="none" stroke="#334155" stroke-width="1" stroke-dasharray="4"/>
        <circle cx="{center}" cy="{center}" r="{radius*0.33}" fill="none" stroke="#334155" stroke-width="1" stroke-dasharray="4"/>

        <!-- Axis Lines -->
        <g stroke="#334155" stroke-width="1">
            {''.join([f'<line x1="{center}" y1="{center}" x2="{grid_points[i].split(",")[0]}" y2="{grid_points[i].split(",")[1]}" />' for i in range(num_axes)])}
        </g>

        <!-- Labels -->
        {''.join(label_elements)}

        <!-- Red Team Area -->
        <polygon points="{' '.join(red_points)}" fill="rgba(239, 68, 68, 0.4)" stroke="#ef4444" stroke-width="2"/>

        <!-- Blue Team Area -->
        <polygon points="{' '.join(blue_points)}" fill="rgba(59, 130, 246, 0.4)" stroke="#3b82f6" stroke-width="2"/>
    </svg>
    """
    return svg

def get_suggestion(blue_team, red_team, current_turn_side, blue_stats, red_stats):
    """Simple heuristic to suggest a hero."""
    my_stats = blue_stats if current_turn_side == 'Blue' else red_stats
    weakest_stat = min(my_stats, key=my_stats.get)

    picked = set(blue_team + red_team)
    candidates = []

    for name, data in hero_data.items():
        if name not in picked:
            score = data['stats'].get(weakest_stat, 0)
            candidates.append((name, score))

    candidates.sort(key=lambda x: x[1], reverse=True)

    if candidates:
        best_hero = candidates[0][0]
        return {
            "team": current_turn_side,
            "reason": f"Team lacks {weakest_stat}. This hero provides strong {weakest_stat}.",
            "heroes": [best_hero]
        }
    return None

# --- ROUTES ---

@app.route('/')
def home():
    # Renders the index.html template provided in the previous step
    return render_template('index.html')

@app.route('/api/heroes', methods=['GET'])
def get_heroes():
    return jsonify(hero_data)

@app.route('/api/analyze', methods=['POST'])
def analyze():
    data = request.json
    blue_team = data.get('blue_team', [])
    red_team = data.get('red_team', [])

    # 1. Calculate WEIGHTED Stats (Using the new Highest Hero logic)
    blue_stats = calculate_weighted_stats(blue_team)
    red_stats = calculate_weighted_stats(red_team)

    # 2. Generate Chart
    svg_chart = generate_radar_chart(blue_stats, red_stats)

    # 3. Calculate Win Probability
    blue_power = sum(blue_stats.values())
    red_power = sum(red_stats.values())
    total_power = blue_power + red_power

    if total_power == 0:
        blue_prob = 50
        red_prob = 50
    else:
        blue_prob = int((blue_power / total_power) * 100)
        red_prob = 100 - blue_prob

    # 4. Determine Advantages
    blue_adv = []
    red_adv = []

    for stat in blue_stats:
        if blue_stats[stat] > red_stats[stat]:
            blue_adv.append(stat)
        elif red_stats[stat] > blue_stats[stat]:
            red_adv.append(stat)

    # 5. Get Suggestion
    turn_side = 'Blue' if len(blue_team) <= len(red_team) else 'Red'
    suggestion = get_suggestion(blue_team, red_team, turn_side, blue_stats, red_stats)

    return jsonify({
        "svg_chart": svg_chart,
        "blue_prob": blue_prob,
        "red_prob": red_prob,
        "blue_adv": blue_adv,
        "red_adv": red_adv,
        "suggestion": suggestion
    })

@app.route('/api/tournament-stats', methods=['GET'])
def get_tournament_stats():
    # Get Filter Parameters
    tournament_filter = request.args.get('tournament', 'All')
    map_filter = request.args.get('map', 'All')
    team_filter = request.args.get('team', 'All')
    side_filter = request.args.get('side', 'Blue')
    sort_by = request.args.get('sort', 'Total Picks')

    # Get DataFrame from Analyzer
    df = analyzer.get_hero_summary(
        side=side_filter,
        tournament_filter=tournament_filter,
        map_filter=map_filter,
        team_filter=team_filter
    )

    # Default sort descending (highest first)
    if not df.empty and sort_by in df.columns:
        df = df.sort_values(by=sort_by, ascending=False)

    # Convert to JSON
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

# --- NEW ROUTE FOR TOURNAMENT DRAFT VIEWER ---
@app.route('/api/matches', methods=['GET'])
def get_matches():
    """
    Returns the list of matches with draft data for the viewer.
    Fetches data from the DataAnalyzer instance.
    """
    # Access the raw matches processed by data_analyzer.py
    raw_matches = analyzer.matches

    matches_list = []

    for idx, match in enumerate(raw_matches):
        # Flatten the data structure for easier consumption by the frontend
        match_data = {
            "id": idx + 1,
            "tournament": match.get('tournament'),
            "map": match.get('map'),
            "blue_team": match.get('blue_team'),
            "red_team": match.get('red_team'),
            "blue_result": match.get('blue_result'),
            "red_result": match.get('red_result'),

            # Blue Side Draft
            "blue_bans_p1": match.get('blue_bans', {}).get('p1', []),
            "blue_bans_p2": match.get('blue_bans', {}).get('p2', []),
            "blue_picks_p1": match.get('blue_picks', {}).get('p1', []),
            "blue_picks_p2": match.get('blue_picks', {}).get('p2', []),

            # Red Side Draft
            "red_bans_p1": match.get('red_bans', {}).get('p1', []),
            "red_bans_p2": match.get('red_bans', {}).get('p2', []),
            "red_picks_p1": match.get('red_picks', {}).get('p1', []),
            "red_picks_p2": match.get('red_picks', {}).get('p2', []),
        }
        matches_list.append(match_data)

    return jsonify(matches_list)

load_data()

if __name__ == '__main__':
    # Run on port 8000
    app.run(debug=False, port=8000)