# 🚀 Setup Instructions

This guide will help you set up automatic updates for your GitHub starred repositories.

## 📋 Prerequisites

- Python 3.11 or higher
- GitHub account with starred repositories
- GitHub Personal Access Token

## 🔧 Initial Setup

### 1. Clone this repository

```bash
git clone <your-repo-url>
cd stars-repos
```

### 2. Install Python dependencies

```bash
pip install -r requirements.txt
```

### 3. Configure your GitHub token (Local Development)

Create a `.env` file in the root directory with your GitHub token:

```
GITHUB_TOKEN=your_github_token_here
```

Replace `your_github_token_here` with your actual GitHub Personal Access Token.

⚠️ **IMPORTANT**: The `.env` file is in `.gitignore` and will NOT be pushed to GitHub for security.

### 4. Test locally

Run the script manually to verify it works:

```bash
python update_starred_repos.py
```

This will generate:
- `README.md` - Beautiful sortable table of all starred repos
- `starred_repos.xlsx` - Excel file with the same data

## 🤖 Automated Updates with GitHub Actions

To enable automatic daily updates, you need to add your GitHub token as a repository secret.

### Step 1: Add GitHub Secret

1. Go to your repository on GitHub
2. Click **Settings** → **Secrets and variables** → **Actions**
3. Click **New repository secret**
4. Name: `STARS_TOKEN`
5. Value: `your_github_token_here` (paste your actual token)
6. Click **Add secret**

### Step 2: Enable GitHub Actions

1. Go to the **Actions** tab in your repository
2. If prompted, click **I understand my workflows, go ahead and enable them**
3. The workflow will now run:
   - **Daily at 00:00 UTC** (midnight)
   - **On every push** to main/master/claude branches
   - **Manually** via the "Run workflow" button

### Step 3: Manual Trigger (Optional)

To run the workflow immediately:

1. Go to **Actions** tab
2. Select **Update Starred Repositories** workflow
3. Click **Run workflow** button
4. Select your branch
5. Click **Run workflow**

## 📊 What Gets Updated Automatically

The GitHub Action will:

1. ✅ Fetch all your starred repositories (handles pagination automatically)
2. ✅ Categorize them by topics and programming language
3. ✅ Generate a beautiful README with:
   - Total statistics (repos, stars, categories)
   - Category distribution chart
   - Sortable HTML table (click headers to sort!)
4. ✅ Create an Excel file with all data
5. ✅ Commit and push changes automatically

## 🎯 Features

### Sortable Table
Click on any column header in the README to sort:
- **Repository Name** - Alphabetically
- **Category** - By category
- **Stars** - By star count
- **Language** - By programming language
- **Last Updated** - By date (newest/oldest)

### Smart Categorization
Repositories are automatically categorized based on:
- Topics (machine-learning, devops, security, etc.)
- Programming language
- Keywords in description

Categories include:
- AI/ML
- Frontend/Backend
- Web Development
- Mobile
- DevOps
- Security
- CLI Tools
- Database
- Blockchain
- And more...

### Excel Export
The `starred_repos.xlsx` file includes:
- All repository data
- Filterable columns
- Frozen header row
- Auto-adjusted column widths
- Easy to import into other tools

## 🔒 Security

- ✅ `.env` file is gitignored (never committed)
- ✅ GitHub token stored as encrypted secret in Actions
- ✅ Token only has read access to public repositories
- ✅ Workflow runs in isolated environment

## 🛠️ Customization

### Change Update Frequency

Edit `.github/workflows/update-stars.yml`:

```yaml
schedule:
  - cron: '0 0 * * *'  # Daily at midnight
  # - cron: '0 */6 * * *'  # Every 6 hours
  # - cron: '0 0 * * 1'  # Weekly on Mondays
```

### Modify Categories

Edit `update_starred_repos.py` in the `categorize_repo()` function to customize category logic.

### Change Table Columns

Edit the `generate_readme()` function to add/remove columns.

## 📝 Manual Updates

You can always run the script manually:

```bash
python update_starred_repos.py
```

Then commit and push:

```bash
git add README.md starred_repos.xlsx
git commit -m "Update starred repositories"
git push
```

## 🐛 Troubleshooting

### Script fails with "GITHUB_TOKEN not found"
- Make sure `.env` file exists with your token
- For GitHub Actions, verify `STARS_TOKEN` secret is set

### GitHub Action fails
- Check if the secret `STARS_TOKEN` is correctly set
- Verify the token has proper permissions
- Check Actions logs for detailed error messages

### No changes detected
- This is normal if you haven't starred new repos
- The workflow will skip committing if nothing changed

## 📞 Support

If you encounter any issues:
1. Check the GitHub Actions logs
2. Run the script locally to see detailed errors
3. Verify your GitHub token is valid and has correct permissions

---

**Enjoy your automatically updated starred repositories list! 🌟**
