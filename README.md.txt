
To push your code to GitHub, follow these step-by-step instructions:
1.	Create a New Repository on GitHub:
o	Go to your GitHub account and click on the "+" sign in the top-right corner to create a new repository. Follow the prompts to set it up 1.
2.	Initialize a Git Repository Locally:
o	If you haven't already, create a new directory on your local machine where you'll store your code. Open a terminal or command prompt, navigate to that directory, and run the following command to initialize a new Git repository:

```bash
rmdir /s /q .git
```

```bash
git init
```

•	Place the files you want to include in your repository in the directory you just created. Then, use the following command to stage them for commit:

```bash
git add .
```

This stages all files in the current directory for commit 1.
3.	Commit Your Changes:
o	After adding files, you need to commit them. A commit is like taking a snapshot of your code at a particular point in time, along with a brief message explaining the changes. Use the following command:

```bash
git commit -m "Your commit message here"
```

•	Branches allow you to work on different versions of your code simultaneously. To check all branch, use the following command:

```bash
git branch
```

This command creates a new branch with the specified name 1.
4.	Switch to the New Branch:
o	To start working on the new branch, use the following command:

```bash
git checkout -b main
```

This command switches your working environment to the specified branch 1.
5.	Make Changes and Commit to the New Branch:
o	Make your code changes and commit them as before using git add . and git commit -m "Your commit message here" 1.
6.	Link Your Local Repository to Your GitHub Repository:
o	Go back to your GitHub repository online. You'll find a section called "Quick setup" with a URL. Copy that URL 1.
o	In your terminal, use the following command to tell Git to connect your local repository to the one on GitHub:

```bash
git remote add origin <paste-the-copied-URL-here>
```

•	Finally, use the following command to send your committed changes to the GitHub repository:

```bash
git push -u origin main
```

To change the remote origin URL in Git, you can follow these steps:
1.	Remove the Existing Remote Origin: Before changing the remote origin URL, it's a good practice to remove the existing remote origin to avoid any conflicts. Use the following command to remove the existing remote named "origin":

```bash
git remote rm origin
```


This command removes the remote named "origin" from your local Git configuration 1.
2.	Add the New Remote Origin: After removing the old remote, you can add the new remote origin with the updated URL. Use the following command to add a new remote named "origin" with the new URL:

```bash
git remote add origin <new-remote-url>
```

Replace <new-remote-url> with the URL of your new remote repository 1.
3.	Verify the Remote URL: To ensure that the remote URL has been updated correctly, you can use the following command to list all remotes for your repository:

```bash
git remote -v
```

This command will display the URLs for all remotes, including the new "origin" URL you just added 4.
4.	Push to the New Remote: After changing the remote origin, you can push your changes to the new remote repository using the following command:

```bash
git push -u origin <branch-name>
```

Replace <branch-name> with the name of the branch you want to push. The -u flag sets the upstream for the branch, so future pushes can be done with just git push 4.
Changing the remote origin URL is a common task when you need to migrate your repository to a different hosting service or when the repository URL has changed. It's important to communicate this change with your collaborators if you're working on a shared project to ensure everyone is using the correct remote URL



This command pushes the code from your local branch to the remote repository (on GitHub). Replace <branch-name> with the name of your branch 1.
Remember, whenever you make new changes, you'll need to repeat steps 3, 4, and 10 to push those changes to your repository. Additionally, if you switch back to the "master" branch (or any other branch), you can continue working on it and commit changes separately from your other branch 1.



The error message you're encountering, ! [rejected] main -> main (fetch first), indicates that the remote repository contains changes that you do not have locally. This situation typically arises when someone else has pushed to the same branch after your last pull or fetch. To resolve this issue, you need to fetch and merge the remote changes into your local branch before you can push your changes. Here's how you can do it:

Fetch the Remote Changes: This step retrieves the latest changes from the remote repository without merging them into your local branch.
```bash
git fetch origin
```

Merge the Remote Changes: After fetching the changes, you need to merge them into your local branch. This step combines the changes from the remote branch with your local changes.
```bash
git merge origin/main
```

Alternatively, you can use git pull to fetch and merge in one step:
```bash
git pull origin main
```

Resolve Any Merge Conflicts: If there are conflicts between your local changes and the changes on the remote, you'll need to resolve these conflicts manually. Open the files with conflicts, decide which changes to keep, and then mark the conflicts as resolved.
Commit the Merged Changes: Once all conflicts are resolved, commit the merged changes to your local branch.
```bash
git commit -m "Merged remote changes"
```

Push Your Changes: Now that your local branch is up to date with the remote, you can push your changes.
```bash
git push origin main
```

It's important to note that using git push -f or --force to forcefully push your changes is generally not recommended, as it can overwrite the remote history and potentially cause issues for other collaborators. This approach should only be used in specific scenarios where you are certain it won't disrupt the work of others 14.

By following these steps, you should be able to resolve the ! [rejected] main -> main (fetch first) error and successfully push your changes to the remote repository.
