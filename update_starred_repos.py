#!/usr/bin/env python3
"""
GitHub Starred Repositories Tracker
Fetches all starred repositories and generates README.md, Excel, and Interactive HTML
"""

import os
import sys
import json
import requests
from datetime import datetime
from typing import List, Dict, Any
from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

# Load environment variables
load_dotenv()

GITHUB_TOKEN = os.getenv('GITHUB_TOKEN')
if not GITHUB_TOKEN:
    print("Error: GITHUB_TOKEN not found in .env file")
    sys.exit(1)

HEADERS = {
    'Authorization': f'token {GITHUB_TOKEN}',
    'Accept': 'application/vnd.github.v3+json'
}

def fetch_all_starred_repos() -> List[Dict[str, Any]]:
    """Fetch all starred repositories with pagination"""
    print("Fetching starred repositories...")
    all_repos = []
    page = 1
    per_page = 100

    while True:
        url = f'https://api.github.com/user/starred?per_page={per_page}&page={page}'
        response = requests.get(url, headers=HEADERS)

        if response.status_code != 200:
            print(f"Error: GitHub API returned status code {response.status_code}")
            print(response.text)
            sys.exit(1)

        repos = response.json()
        if not repos:
            break

        all_repos.extend(repos)
        print(f"  Fetched page {page} ({len(repos)} repos)")
        page += 1

    print(f"Total repositories fetched: {len(all_repos)}")
    return all_repos

def categorize_repo(repo: Dict[str, Any]) -> str:
    """Categorize repository based on language and topics"""
    language = repo.get('language', '')
    topics = repo.get('topics', [])

    # Define category mappings
    web_langs = ['JavaScript', 'TypeScript', 'HTML', 'CSS', 'Vue', 'React']
    backend_langs = ['Python', 'Java', 'Go', 'Ruby', 'PHP', 'C#', 'Rust', 'Kotlin']
    mobile_langs = ['Swift', 'Objective-C', 'Dart', 'Kotlin', 'Java']
    data_langs = ['Python', 'R', 'Julia', 'Jupyter Notebook']

    # Check topics first for more specific categorization
    topic_str = ' '.join(topics).lower()

    if any(t in topic_str for t in ['machine-learning', 'deep-learning', 'ai', 'ml', 'neural-network', 'data-science']):
        return 'AI/ML'
    elif any(t in topic_str for t in ['devops', 'docker', 'kubernetes', 'cicd', 'ci-cd']):
        return 'DevOps'
    elif any(t in topic_str for t in ['security', 'cybersecurity', 'penetration-testing']):
        return 'Security'
    elif any(t in topic_str for t in ['frontend', 'react', 'vue', 'angular', 'web']):
        return 'Frontend'
    elif any(t in topic_str for t in ['backend', 'api', 'server']):
        return 'Backend'
    elif any(t in topic_str for t in ['mobile', 'android', 'ios', 'flutter']):
        return 'Mobile'
    elif any(t in topic_str for t in ['cli', 'command-line', 'terminal']):
        return 'CLI Tools'
    elif any(t in topic_str for t in ['database', 'sql', 'nosql']):
        return 'Database'
    elif any(t in topic_str for t in ['blockchain', 'cryptocurrency', 'web3']):
        return 'Blockchain'

    # Fallback to language-based categorization
    if language in web_langs:
        return 'Web Development'
    elif language in backend_langs:
        return 'Backend'
    elif language in mobile_langs:
        return 'Mobile'
    elif language == 'Python' and any(t in topic_str for t in ['data', 'analysis', 'science']):
        return 'Data Science'
    elif language:
        return language
    else:
        return 'Other'

def truncate_description(description: str, max_sentences: int = 3) -> str:
    """Truncate description to max sentences"""
    if not description:
        return "No description available"

    # Simple sentence splitting
    sentences = description.replace('! ', '!|').replace('? ', '?|').replace('. ', '.|').split('|')
    truncated = ' '.join(sentences[:max_sentences]).strip()

    # Limit to reasonable length
    if len(truncated) > 200:
        truncated = truncated[:197] + '...'

    return truncated

def format_date(date_str: str) -> str:
    """Format ISO date to DD-MM-YYYY HH:MM"""
    if not date_str:
        return "N/A"

    try:
        dt = datetime.strptime(date_str, '%Y-%m-%dT%H:%M:%SZ')
        return dt.strftime('%d-%m-%Y %H:%M')
    except:
        return date_str

def process_repos(repos: List[Dict[str, Any]]) -> List[Dict[str, str]]:
    """Process repositories into structured data"""
    print("Processing repositories...")
    processed = []

    for repo in repos:
        processed.append({
            'name': repo['full_name'],
            'url': repo['html_url'],
            'category': categorize_repo(repo),
            'description': truncate_description(repo.get('description', '')),
            'stars': repo.get('stargazers_count', 0),
            'language': repo.get('language', 'N/A'),
            'last_updated': format_date(repo.get('updated_at', '')),
            'created_at': format_date(repo.get('created_at', '')),
            'topics': ', '.join(repo.get('topics', [])[:5])  # First 5 topics
        })

    # Sort by stars by default (descending)
    processed.sort(key=lambda x: x['stars'], reverse=True)

    print(f"Processed {len(processed)} repositories")
    return processed

def generate_interactive_html(repos: List[Dict[str, str]]) -> None:
    """Generate interactive HTML file with sortable table and search"""
    print("Generating interactive HTML...")

    last_updated = datetime.utcnow().strftime('%d-%m-%Y %H:%M UTC')
    total_repos = len(repos)
    total_stars = sum(r['stars'] for r in repos)

    categories = {}
    for repo in repos:
        cat = repo['category']
        categories[cat] = categories.get(cat, 0) + 1

    html_content = f'''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>My GitHub Starred Repositories</title>
    <style>
        * {{ box-sizing: border-box; }}
        body {{
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Helvetica, Arial, sans-serif;
            max-width: 1600px;
            margin: 0 auto;
            padding: 20px;
            background: #0d1117;
            color: #c9d1d9;
            line-height: 1.6;
        }}
        h1 {{
            color: #58a6ff;
            text-align: center;
            margin-bottom: 10px;
        }}
        .subtitle {{
            text-align: center;
            color: #8b949e;
            margin-bottom: 30px;
        }}
        .stats {{
            background: #161b22;
            padding: 20px;
            border-radius: 6px;
            margin: 20px 0;
            border: 1px solid #30363d;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
            background: #161b22;
            border-radius: 6px;
            overflow: hidden;
            border: 1px solid #30363d;
        }}
        th {{
            background: #21262d;
            padding: 12px;
            text-align: left;
            font-weight: 600;
            cursor: pointer;
            user-select: none;
            position: sticky;
            top: 0;
            z-index: 10;
            white-space: nowrap;
        }}
        th:hover {{
            background: #30363d;
        }}
        td {{
            padding: 10px 12px;
            border-top: 1px solid #21262d;
            word-wrap: break-word;
        }}
        tr:hover {{
            background: #0d1117;
        }}
        a {{
            color: #58a6ff;
            text-decoration: none;
        }}
        a:hover {{
            text-decoration: underline;
        }}
        .stats-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 15px;
        }}
        .stat-box {{
            background: #0d1117;
            padding: 15px;
            border-radius: 6px;
            border: 1px solid #30363d;
        }}
        .stat-number {{
            font-size: 24px;
            font-weight: bold;
            color: #58a6ff;
        }}
        .stat-label {{
            color: #8b949e;
            font-size: 14px;
        }}
        .controls {{
            margin: 20px 0;
            text-align: center;
        }}
        input[type="text"] {{
            padding: 10px 15px;
            border: 1px solid #30363d;
            background: #0d1117;
            color: #c9d1d9;
            border-radius: 6px;
            width: 100%;
            max-width: 500px;
            font-size: 14px;
        }}
        input[type="text"]:focus {{
            outline: none;
            border-color: #58a6ff;
        }}
        .sort-indicator {{
            font-size: 0.8em;
            margin-left: 5px;
        }}
        /* Column widths */
        th:nth-child(1), td:nth-child(1) {{ width: 20%; }} /* Name */
        th:nth-child(2), td:nth-child(2) {{ width: 12%; text-align: center; }} /* Category */
        th:nth-child(3), td:nth-child(3) {{ width: 8%; text-align: center; }} /* Stars */
        th:nth-child(4), td:nth-child(4) {{ width: 10%; text-align: center; }} /* Language */
        th:nth-child(5), td:nth-child(5) {{ width: 38%; }} /* Description */
        th:nth-child(6), td:nth-child(6) {{ width: 12%; text-align: center; font-size: 0.85em; }} /* Updated */
    </style>
</head>
<body>
    <h1>🌟 My GitHub Starred Repositories</h1>
    <p class="subtitle">Automatically updated tracker of all starred repositories</p>

    <div class="stats">
        <div class="stats-grid">
            <div class="stat-box">
                <div class="stat-number">{total_repos:,}</div>
                <div class="stat-label">Total Repositories</div>
            </div>
            <div class="stat-box">
                <div class="stat-number">{total_stars:,}</div>
                <div class="stat-label">Total Stars Given</div>
            </div>
            <div class="stat-box">
                <div class="stat-number">{len(categories)}</div>
                <div class="stat-label">Categories</div>
            </div>
            <div class="stat-box">
                <div class="stat-number">{last_updated}</div>
                <div class="stat-label">Last Updated</div>
            </div>
        </div>
    </div>

    <div class="controls">
        <input type="text" id="searchInput" placeholder="🔍 Search repositories by name, category, language, or description..." onkeyup="searchTable()">
    </div>

    <table id="repoTable">
        <thead>
            <tr>
                <th onclick="sortTable(0)">Repository Name <span id="sort0" class="sort-indicator">🔽</span></th>
                <th onclick="sortTable(1)">Category <span id="sort1" class="sort-indicator">🔽</span></th>
                <th onclick="sortTable(2)">Stars ⭐ <span id="sort2" class="sort-indicator">🔽</span></th>
                <th onclick="sortTable(3)">Language <span id="sort3" class="sort-indicator">🔽</span></th>
                <th>Description</th>
                <th onclick="sortTable(5)">Last Updated <span id="sort5" class="sort-indicator">🔽</span></th>
            </tr>
        </thead>
        <tbody>
'''

    # Add repository rows
    for repo in repos:
        html_content += f'''            <tr>
                <td><a href="{repo['url']}" target="_blank">{repo['name']}</a></td>
                <td>{repo['category']}</td>
                <td>{repo['stars']:,}</td>
                <td>{repo['language']}</td>
                <td>{repo['description']}</td>
                <td>{repo['last_updated']}</td>
            </tr>
'''

    html_content += '''        </tbody>
    </table>

    <script>
        let sortDirection = {};

        function sortTable(columnIndex) {
            const table = document.getElementById("repoTable");
            const tbody = table.querySelector("tbody");
            const rows = Array.from(tbody.querySelectorAll("tr"));

            // Toggle sort direction
            sortDirection[columnIndex] = !sortDirection[columnIndex];
            const ascending = sortDirection[columnIndex];

            rows.sort((a, b) => {
                let aValue = a.cells[columnIndex].textContent.trim();
                let bValue = b.cells[columnIndex].textContent.trim();

                // Handle numeric sorting for stars
                if (columnIndex === 2) {
                    aValue = parseInt(aValue.replace(/,/g, ''));
                    bValue = parseInt(bValue.replace(/,/g, ''));
                }

                // Handle date sorting
                if (columnIndex === 5) {
                    const parseDate = (dateStr) => {
                        if (dateStr === 'N/A') return new Date(0);
                        const [date, time] = dateStr.split(' ');
                        if (!date) return new Date(0);
                        const [day, month, year] = date.split('-');
                        return new Date(`${year}-${month}-${day}T${time || '00:00'}:00`);
                    };
                    aValue = parseDate(aValue);
                    bValue = parseDate(bValue);
                }

                if (aValue < bValue) return ascending ? -1 : 1;
                if (aValue > bValue) return ascending ? 1 : -1;
                return 0;
            });

            // Re-append sorted rows
            rows.forEach(row => tbody.appendChild(row));

            // Update sort indicators
            document.querySelectorAll('.sort-indicator').forEach(span => {
                span.textContent = '🔽';
            });
            document.getElementById(`sort${columnIndex}`).textContent = ascending ? '🔼' : '🔽';
        }

        function searchTable() {
            const input = document.getElementById("searchInput");
            const filter = input.value.toLowerCase();
            const table = document.getElementById("repoTable");
            const tr = table.getElementsByTagName("tr");

            for (let i = 1; i < tr.length; i++) {
                const row = tr[i];
                const text = row.textContent.toLowerCase();
                row.style.display = text.includes(filter) ? "" : "none";
            }
        }

        // Initialize - sort by stars (already sorted in Python, but show indicator)
        sortDirection[2] = false; // Will toggle to true on first click
    </script>
</body>
</html>
'''

    with open('starred_repos.html', 'w', encoding='utf-8') as f:
        f.write(html_content)

    print("Interactive HTML generated successfully!")

def generate_readme(repos: List[Dict[str, str]]) -> None:
    """Generate README.md with properly formatted table"""
    print("Generating README.md...")

    # Calculate statistics
    total_repos = len(repos)
    total_stars = sum(r['stars'] for r in repos)
    categories = {}
    for repo in repos:
        cat = repo['category']
        categories[cat] = categories.get(cat, 0) + 1

    # Get last updated time
    last_updated = datetime.utcnow().strftime('%d-%m-%Y %H:%M UTC')

    readme_content = f'''# 🌟 My GitHub Starred Repositories

> Automatically updated list of all my starred repositories on GitHub

## 📊 Statistics

- **Total Repositories**: {total_repos:,}
- **Total Stars Given**: {total_stars:,}
- **Categories**: {len(categories)}
- **Last Updated**: {last_updated}

### 📁 Category Distribution

'''

    # Add category statistics
    for cat, count in sorted(categories.items(), key=lambda x: x[1], reverse=True):
        percentage = (count / total_repos) * 100
        readme_content += f"- **{cat}**: {count} repos ({percentage:.1f}%)\n"

    # Show top 50 repos in README to avoid cluttering
    display_limit = 50

    readme_content += f'''
---

## 📥 Interactive Versions

For the full experience with **sortable columns** and **live search functionality**:

- 🌐 **[Interactive HTML Table](./starred_repos.html)** - Download and open in your browser for sorting and searching
- 📊 **[Excel Spreadsheet](./starred_repos.xlsx)** - Full data with filters and sorting

> **Note**: GitHub doesn't support JavaScript in README files, so the table below is static.
> Use the interactive HTML file above for sorting, filtering, and search features!

---

## 🗂️ Top {display_limit} Starred Repositories (by stars)

<table>
<thead>
<tr>
<th align="left">Repository</th>
<th align="center">Category</th>
<th align="center">Stars ⭐</th>
<th align="center">Language</th>
<th align="left">Description</th>
<th align="center">Last Updated</th>
</tr>
</thead>
<tbody>
'''

    # Add repository rows (limited to display_limit)
    for repo in repos[:display_limit]:
        # Truncate description for README display
        desc = repo['description']
        if len(desc) > 100:
            desc = desc[:97] + '...'

        readme_content += f'''<tr>
<td><a href="{repo['url']}">{repo['name']}</a></td>
<td align="center">{repo['category']}</td>
<td align="center">{repo['stars']:,}</td>
<td align="center">{repo['language']}</td>
<td>{desc}</td>
<td align="center"><sub>{repo['last_updated']}</sub></td>
</tr>
'''

    readme_content += f'''</tbody>
</table>

<details>
<summary><b>📋 View All {total_repos} Repositories by Category</b></summary>
<br>

'''

    # Group repositories by category
    by_category = {}
    for repo in repos:
        cat = repo['category']
        if cat not in by_category:
            by_category[cat] = []
        by_category[cat].append(repo)

    # Display repositories grouped by category
    for cat in sorted(by_category.keys()):
        cat_repos = by_category[cat]
        readme_content += f'''### {cat} ({len(cat_repos)} repositories)

'''
        for repo in cat_repos[:20]:  # Show max 20 per category
            desc_short = repo['description'][:100] + ("..." if len(repo['description']) > 100 else "")
            readme_content += f'''- [{repo['name']}]({repo['url']}) ⭐ {repo['stars']:,} - {desc_short}\n'''

        if len(cat_repos) > 20:
            readme_content += f'''\n*...and {len(cat_repos) - 20} more. See [interactive HTML](./starred_repos.html) or [Excel file](./starred_repos.xlsx) for complete list.*\n'''

        readme_content += '\n'

    readme_content += f'''
</details>

---

## 🔄 Automatic Updates

This repository is automatically updated **daily at 00:00 UTC** using GitHub Actions.

**What gets updated:**
- ✅ All {total_repos:,} starred repositories (with pagination support)
- ✅ Star counts and update dates
- ✅ New repositories you star
- ✅ Category statistics and analytics
- ✅ README.md, Excel file, and Interactive HTML

---

## 👥 Want to use this for your own starred repos?

Check out the **[TUTORIAL.md](./TUTORIAL.md)** for complete setup instructions!

---

<div align="center">

**Made with ❤️ using Python and GitHub Actions**

*Last generated: {last_updated}*

[📊 View Interactive Version](./starred_repos.html) • [📥 Download Excel](./starred_repos.xlsx) • [📖 Setup Tutorial](./TUTORIAL.md)

</div>
'''

    with open('README.md', 'w', encoding='utf-8') as f:
        f.write(readme_content)

    print("README.md generated successfully!")

def generate_excel(repos: List[Dict[str, str]]) -> None:
    """Generate Excel file with repository data"""
    print("Generating Excel file...")

    wb = Workbook()
    ws = wb.active
    ws.title = "Starred Repositories"

    # Define headers
    headers = ['Repository Name', 'URL', 'Category', 'Description', 'Stars', 'Language', 'Last Updated', 'Created At', 'Topics']

    # Style headers
    header_fill = PatternFill(start_color='366092', end_color='366092', fill_type='solid')
    header_font = Font(bold=True, color='FFFFFF', size=12)

    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')

    # Add data
    for row_idx, repo in enumerate(repos, 2):
        ws.cell(row=row_idx, column=1, value=repo['name'])
        ws.cell(row=row_idx, column=2, value=repo['url'])
        ws.cell(row=row_idx, column=3, value=repo['category'])
        ws.cell(row=row_idx, column=4, value=repo['description'])
        ws.cell(row=row_idx, column=5, value=repo['stars'])
        ws.cell(row=row_idx, column=6, value=repo['language'])
        ws.cell(row=row_idx, column=7, value=repo['last_updated'])
        ws.cell(row=row_idx, column=8, value=repo['created_at'])
        ws.cell(row=row_idx, column=9, value=repo['topics'])

    # Auto-adjust column widths
    for col in range(1, len(headers) + 1):
        column_letter = get_column_letter(col)
        max_length = 0

        for row in ws[column_letter]:
            try:
                if len(str(row.value)) > max_length:
                    max_length = len(str(row.value))
            except:
                pass

        adjusted_width = min(max_length + 2, 50)  # Max width of 50
        ws.column_dimensions[column_letter].width = adjusted_width

    # Enable filters
    ws.auto_filter.ref = ws.dimensions

    # Freeze header row
    ws.freeze_panes = 'A2'

    # Save workbook
    wb.save('starred_repos.xlsx')
    print("Excel file generated successfully!")

def main():
    """Main execution function"""
    print("=" * 60)
    print("GitHub Starred Repositories Tracker")
    print("=" * 60)

    # Fetch repositories
    repos = fetch_all_starred_repos()

    # Process data
    processed_repos = process_repos(repos)

    # Generate outputs
    generate_readme(processed_repos)
    generate_interactive_html(processed_repos)
    generate_excel(processed_repos)

    print("=" * 60)
    print("✅ All done! Files generated:")
    print("  - README.md")
    print("  - starred_repos.html (interactive)")
    print("  - starred_repos.xlsx")
    print("=" * 60)

if __name__ == '__main__':
    main()
