# Setting Up VS Code to Connect to Your Git Repository

This guide will walk you through the process of setting up Visual Studio Code on your local computer to connect to the workTools Git repository, and how to commit and push changes.

## Prerequisites

Before you begin, make sure you have:

1. [Git](https://git-scm.com/downloads) installed on your computer
2. [Visual Studio Code](https://code.visualstudio.com/download) installed on your computer
3. A GitHub account with access to the `rachelvba/workTools` repository

## Step 1: Install Git and VS Code

### Installing Git

1. Download Git from [git-scm.com/downloads](https://git-scm.com/downloads)
2. Run the installer and follow the installation instructions
3. Open a terminal or command prompt and verify the installation with:
   ```
   git --version
   ```

### Installing Visual Studio Code

1. Download VS Code from [code.visualstudio.com/download](https://code.visualstudio.com/download)
2. Run the installer and follow the installation instructions

## Step 2: Configure Git

Open a terminal or command prompt and set your Git username and email:

```
git config --global user.name "Your Name"
git config --global user.email "your.email@example.com"
```

## Step 3: Authenticate with GitHub

### Generate SSH Key (Recommended Method)

1. Open a terminal or command prompt
2. Generate a new SSH key:
   ```
   ssh-keygen -t ed25519 -C "your.email@example.com"
   ```
3. Start the SSH agent:
   - **Windows**: 
     ```
     eval `ssh-agent -s`
     ```
   - **macOS/Linux**:
     ```
     eval "$(ssh-agent -s)"
     ```
4. Add your SSH key to the agent:
   ```
   ssh-add ~/.ssh/id_ed25519
   ```
5. Copy the SSH public key to your clipboard:
   - **Windows**: 
     ```
     cat ~/.ssh/id_ed25519.pub | clip
     ```
   - **macOS**: 
     ```
     pbcopy < ~/.ssh/id_ed25519.pub
     ```
   - **Linux**: 
     ```
     cat ~/.ssh/id_ed25519.pub | xclip -selection clipboard
     ```

### Add SSH Key to GitHub

1. Go to [GitHub Settings > SSH and GPG keys](https://github.com/settings/keys)
2. Click "New SSH key"
3. Give your key a title (e.g., "Work Laptop")
4. Paste the key from your clipboard into the "Key" field
5. Click "Add SSH key"

## Step 4: Clone the Repository in VS Code

1. Open VS Code
2. Press `Ctrl+Shift+P` (Windows/Linux) or `Cmd+Shift+P` (macOS) to open the Command Palette
3. Type "Git: Clone" and select it
4. Enter the SSH URL of the repository:
   ```
   git@github.com:rachelvba/workTools.git
   ```
   Alternatively, use HTTPS URL if you prefer password authentication:
   ```
   https://github.com/rachelvba/workTools.git
   ```
5. Select a local folder to store the repository
6. When prompted, sign in to GitHub

## Step 5: Working with the Repository

### Opening the Repository

1. In VS Code, go to "File" > "Open Folder"
2. Navigate to and select the cloned repository folder

### Making Changes

1. Edit files in VS Code as needed
2. Save your changes with `Ctrl+S` (Windows/Linux) or `Cmd+S` (macOS)

### Committing Changes

1. Click on the Source Control icon in the Activity Bar (or press `Ctrl+Shift+G`)
2. Review your changes
3. Enter a commit message in the text field at the top of the Source Control panel
4. Click the checkmark (✓) icon to commit changes

### Pushing Changes

1. After committing, click the "..." menu in the Source Control panel
2. Select "Push"
3. Alternatively, click on the sync icon at the bottom of the VS Code window

## Step 6: Pulling Latest Changes

To get the latest updates from the repository:

1. Click on the "..." menu in the Source Control panel
2. Select "Pull"
3. Alternatively, click on the sync icon at the bottom of the VS Code window

## Step 7: Using VS Code Extensions for Finance Work

Install these helpful extensions for working with financial code:

1. Click on the Extensions icon in the Activity Bar
2. Search for and install:
   - "Excel Viewer" - for viewing Excel files
   - "Python" - for Python development
   - "Jupyter" - for working with Jupyter notebooks
   - "VSCode VBA" - for VBA syntax highlighting
   - "GitHub Copilot" - for AI-assisted coding

## Step 8: Setting Up GitHub Copilot

1. Install the GitHub Copilot extension from the VS Code Marketplace
2. Sign in to GitHub when prompted
3. If you have a GitHub Copilot subscription, it will activate automatically
4. Follow the GitHub Copilot Agent instructions in the repository's README.md file

## Troubleshooting

### Authentication Issues

If you have trouble authenticating:

1. Ensure your SSH key is added to GitHub properly
2. Try using a personal access token instead for HTTPS authentication
   - Create a token at [GitHub Settings > Developer settings > Personal access tokens](https://github.com/settings/tokens)
   - Use this token as your password when prompted

### Connection Issues

If VS Code can't connect to the repository:

1. Check your internet connection
2. Verify you have the correct repository URL
3. Ensure you have proper access permissions to the repository

### Git Command Issues

If Git commands fail:

1. Ensure Git is properly installed and in your PATH
2. Restart VS Code
3. Try running Git commands from the terminal to identify specific errors

## Additional Resources

- [VS Code Documentation](https://code.visualstudio.com/docs)
- [Git Documentation](https://git-scm.com/doc)
- [GitHub Documentation](https://docs.github.com/en)

For additional help, contact your team's technical support or repository administrator.
