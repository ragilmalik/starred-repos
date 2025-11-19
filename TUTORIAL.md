# 📖 Tutorial: Set Up Your Own GitHub Starred Repositories Tracker

This tutorial will guide you through setting up this automated tracker for your own GitHub starred repositories.

## 📋 Table of Contents

1. [Prerequisites](#prerequisites)
2. [Getting Started](#getting-started)
3. [Local Setup](#local-setup)
4. [GitHub Token Setup](#github-token-setup)
5. [Automation with GitHub Actions](#automation-with-github-actions)
6. [Customization Options](#customization-options)
7. [Troubleshooting](#troubleshooting)

---

## Prerequisites

Before you begin, make sure you have:

- ✅ A GitHub account
- ✅ Python 3.11 or higher installed
- ✅ Git installed on your computer
- ✅ Basic knowledge of terminal/command line

---

## Getting Started

### Step 1: Fork or Clone This Repository

**Option A: Fork this repository (Recommended)**

1. Click the "Fork" button at the top-right of this repository
2. This creates a copy under your GitHub account
3. Clone your fork:
   ```bash
   git clone https://github.com/YOUR_USERNAME/stars-repos.git
   cd stars-repos
   ```

**Option B: Create a new repository**

1. Create a new repository on GitHub
2. Clone the repository locally
3. Copy all files from this project to your repository

### Step 2: Install Dependencies

```bash
pip install -r requirements.txt
```

This installs:
- `requests` - For GitHub API calls
- `python-dotenv` - For environment variable management
- `openpyxl` - For Excel file generation
- `Jinja2` - For template rendering

---

## Local Setup

### Step 3: Create GitHub Personal Access Token

1. Go to GitHub Settings: https://github.com/settings/tokens
2. Click **"Generate new token"** → **"Generate new token (classic)"**
3. Give it a descriptive name like "Starred Repos Tracker"
4. Select the following scopes:
   - ✅ `repo` (Full control of private repositories)
   - ✅ `read:user` (Read ALL user profile data)
5. Set expiration (30 days recommended for testing, no expiration for permanent use)
6. Click **"Generate token"**
7. **Copy the token** (you won't see it again!)

### Step 4: Configure Local Environment

Create a `.env` file in the root directory:

```bash
touch .env
```

Add your GitHub token to `.env`:

```
GITHUB_TOKEN=your_github_token_here
```

Replace `your_github_token_here` with the token you just created.

⚠️ **IMPORTANT**: The `.env` file is gitignored and will NOT be pushed to GitHub for security.

### Step 5: Test Locally

Run the script to generate your starred repositories:

```bash
python update_starred_repos.py
```

This will create:
- `README.md` - Overview with top 50 repos
- `starred_repos.html` - Interactive sortable table with search
- `starred_repos.xlsx` - Excel file with all data

---

## GitHub Token Setup

### Step 6: Add Token as GitHub Secret

For automatic updates via GitHub Actions, you need to add your token as a repository secret:

1. Go to your repository on GitHub
2. Click **Settings** → **Secrets and variables** → **Actions**
3. Click **"New repository secret"**
4. Name: `STARS_TOKEN`
5. Value: Paste your GitHub Personal Access Token
6. Click **"Add secret"**

---

## Automation with GitHub Actions

### Step 7: Enable GitHub Actions

The `.github/workflows/update-stars.yml` file is already configured to run automatically.

1. Go to the **Actions** tab in your repository
2. If prompted, click **"I understand my workflows, go ahead and enable them"**
3. The workflow will now run:
   - **Daily at 00:00 UTC** (midnight)
   - **On every push** to main/master/claude branches
   - **Manually** via the "Run workflow" button

### Step 8: Test the Workflow

To manually trigger the workflow:

1. Go to **Actions** tab
2. Click **"Update Starred Repositories"** workflow
3. Click **"Run workflow"** button
4. Select your branch (usually `main` or `master`)
5. Click **"Run workflow"**

Wait a few minutes and check:
- The workflow completes successfully (green checkmark)
- Your README.md is updated
- Files `starred_repos.xlsx` and `starred_repos.html` are generated
- A new commit appears in your repository

---

## Customization Options

### Change Update Frequency

Edit `.github/workflows/update-stars.yml`:

```yaml
on:
  schedule:
    - cron: '0 0 * * *'  # Daily at midnight UTC
    # - cron: '0 */6 * * *'  # Every 6 hours
    # - cron: '0 0 * * 1'  # Weekly on Mondays
    # - cron: '0 9 * * *'  # Daily at 9 AM UTC
```

### Modify Categories

Edit `update_starred_repos.py` in the `categorize_repo()` function (lines 58-103):

```python
def categorize_repo(repo: Dict[str, Any]) -> str:
    """Categorize repository based on language and topics"""
    # Add your own category logic here
    topic_str = ' '.join(repo.get('topics', [])).lower()

    if 'your-custom-topic' in topic_str:
        return 'Your Custom Category'
    # ... rest of the categorization logic
```

### Change Table Display Limit

Edit `update_starred_repos.py` in the `generate_readme()` function (line 461):

```python
display_limit = 50  # Change this to show more/fewer repos in README
```

### Customize HTML Styling

Edit the `<style>` section in the `generate_interactive_html()` function (lines 174-290) to change colors, fonts, layout, etc.

### Add More Statistics

Edit the statistics section in both `generate_readme()` and `generate_interactive_html()` functions to add:
- Top languages count
- Average stars per repository
- Most active categories
- Custom metrics

---

## Troubleshooting

### Issue: "GITHUB_TOKEN not found in .env file"

**Solution**: Make sure you created the `.env` file and added your token:
```
GITHUB_TOKEN=your_actual_token_here
```

### Issue: GitHub Action fails with "Error 401 Unauthorized"

**Solution**:
- Check that `STARS_TOKEN` secret is set in repository settings
- Verify your token hasn't expired
- Ensure the token has `repo` and `read:user` scopes

### Issue: No files are generated

**Solution**:
- Check Python is installed: `python --version`
- Ensure dependencies are installed: `pip install -r requirements.txt`
- Check for error messages in the script output

### Issue: Table not sorting in README on GitHub

**This is expected behavior!** GitHub doesn't support JavaScript in README files for security reasons.

**Solutions**:
- Download and open `starred_repos.html` in your browser for full sorting functionality
- Use the Excel file for advanced filtering and sorting
- The README table is intentionally static and shows top 50 repos only

### Issue: Script is slow with many starred repos

**This is normal!** The GitHub API rate limits requests.

**Optimizations**:
- The script uses pagination (100 repos per request)
- Consider caching results if running frequently locally
- GitHub Actions runs on remote servers with good bandwidth

### Issue: Some repositories are miscategorized

**Solution**:
Edit the `categorize_repo()` function in `update_starred_repos.py` to refine the logic:

```python
def categorize_repo(repo: Dict[str, Any]) -> str:
    # Adjust the topic keywords
    # Add more specific categories
    # Change priority order
```

---

## Advanced Configuration

### Running on a Schedule Locally (Optional)

If you want to run this locally on a schedule:

**On macOS/Linux** (using cron):

1. Open crontab:
   ```bash
   crontab -e
   ```

2. Add a daily job at 9 AM:
   ```
   0 9 * * * cd /path/to/stars-repos && /usr/bin/python3 update_starred_repos.py
   ```

**On Windows** (using Task Scheduler):

1. Open Task Scheduler
2. Create a new task
3. Set trigger to daily
4. Set action to run `python.exe` with argument `C:\path\to\stars-repos\update_starred_repos.py`

### Using with Private Repositories

The default token scopes (`repo` and `read:user`) already include access to private repositories.

If you only want public repos, create a token with just:
- `public_repo` scope
- `read:user` scope

### Multiple GitHub Accounts

To track starred repos from multiple accounts:

1. Create separate `.env` files: `.env.account1`, `.env.account2`
2. Modify the script to accept environment file as argument
3. Run separately: `python update_starred_repos.py --env .env.account1`

---

## What Gets Generated

### README.md
- Statistics (total repos, stars, categories)
- Category distribution chart
- Top 50 repositories table
- All repos grouped by category (collapsible)
- Links to interactive versions

### starred_repos.html
- **Fully interactive table** with JavaScript
- **Click headers to sort** (repository name, category, stars, language, date)
- **Live search** functionality
- **Responsive design** for mobile and desktop
- **Dark theme** matching GitHub's style
- Fixed column widths to prevent cutoff

### starred_repos.xlsx
- All repository data in Excel format
- Filterable columns
- Frozen header row
- Auto-adjusted column widths
- Easy to import into other tools

---

## Security Best Practices

1. ✅ **Never commit your `.env` file** - It's in `.gitignore`
2. ✅ **Use GitHub Secrets** for automation - Encrypted and secure
3. ✅ **Set token expiration** - Regularly rotate tokens
4. ✅ **Minimum required scopes** - Only grant necessary permissions
5. ✅ **Review repository secrets** - Periodically audit access

---

## Need Help?

If you encounter issues:

1. **Check the troubleshooting section above**
2. **Review GitHub Actions logs** for detailed error messages
3. **Test locally first** before relying on automation
4. **Verify your token permissions** on GitHub settings

---

## Contributing

Found a bug or want to add a feature? Contributions are welcome!

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

---

## License

This project is open source. Feel free to use, modify, and distribute as needed!

---

<div align="center">

**Happy tracking! 🌟**

*Made with ❤️ by the community*

</div>
