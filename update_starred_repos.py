#!/usr/bin/env python3
"""
GitHub Starred Repositories Tracker
Fetches all starred repositories and generates README.md and Excel file
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

def generate_readme(repos: List[Dict[str, str]]) -> None:
    """Generate README.md with sortable HTML table"""
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

- **Total Repositories**: {total_repos}
- **Total Stars Given**: {total_stars:,}
- **Categories**: {len(categories)}
- **Last Updated**: {last_updated}

### 📁 Category Distribution

'''

    # Add category statistics
    for cat, count in sorted(categories.items(), key=lambda x: x[1], reverse=True):
        percentage = (count / total_repos) * 100
        readme_content += f"- **{cat}**: {count} repos ({percentage:.1f}%)\n"

    readme_content += '''
---

## 🗂️ All Starred Repositories

> Click on any column header to sort the table!

<div align="center">

<table id="repoTable">
<thead>
<tr>
<th onclick="sortTable(0)" style="cursor: pointer;">Repository Name 🔽</th>
<th onclick="sortTable(1)" style="cursor: pointer;">Category 🔽</th>
<th onclick="sortTable(2)" style="cursor: pointer;">Stars ⭐ 🔽</th>
<th onclick="sortTable(3)" style="cursor: pointer;">Language 🔽</th>
<th>Description</th>
<th onclick="sortTable(5)" style="cursor: pointer;">Last Updated 🔽</th>
</tr>
</thead>
<tbody>
'''

    # Add repository rows
    for repo in repos:
        readme_content += f'''<tr>
<td><a href="{repo['url']}">{repo['name']}</a></td>
<td>{repo['category']}</td>
<td>{repo['stars']:,}</td>
<td>{repo['language']}</td>
<td>{repo['description']}</td>
<td>{repo['last_updated']}</td>
</tr>
'''

    readme_content += '''</tbody>
</table>

</div>

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
            // Convert DD-MM-YYYY HH:MM to comparable format
            const parseDate = (dateStr) => {
                if (dateStr === 'N/A') return new Date(0);
                const [date, time] = dateStr.split(' ');
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

    // Update header indicators
    const headers = table.querySelectorAll("th[onclick]");
    headers.forEach((header, index) => {
        const text = header.textContent.replace(' 🔽', '').replace(' 🔼', '');
        if (index === columnIndex) {
            header.textContent = text + (ascending ? ' 🔼' : ' 🔽');
        } else {
            header.textContent = text.replace(' 🔼', ' 🔽');
        }
    });
}
</script>

---

## 🔄 Automatic Updates

This repository is automatically updated daily using GitHub Actions. The workflow:

1. Fetches all starred repositories via GitHub API
2. Categorizes them by topics and language
3. Generates this README with sortable table
4. Creates an Excel file (`starred_repos.xlsx`) with the same data
5. Commits and pushes changes automatically

## 📥 Download Excel

Download the complete list: [starred_repos.xlsx](./starred_repos.xlsx)

---

<div align="center">

**Made with ❤️ using Python and GitHub Actions**

*Last generated: {last_updated}*

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
    generate_excel(processed_repos)

    print("=" * 60)
    print("✅ All done! Files generated:")
    print("  - README.md")
    print("  - starred_repos.xlsx")
    print("=" * 60)

if __name__ == '__main__':
    main()
