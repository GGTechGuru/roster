# roster

This repository serves as a programming project to assess your understanding of certain Python concepts and give you an opportunity to demonstrate your Python proficiency by crafting a solution class and sharing your design decisions.

## the problem

Mrs. Jones, a teacher at Billings Elementary School, has been asked by the administration to track her students' grades in a roster stored in an Excel file. Unfortunatly, Mrs. Jones had a bad experience with Excel as a child and as solemnly sworn never to use Excel again. Your task is to create a Python class for reading and manipulating a provided Excel file so Mrs. Jones doesn't haven't to use Excel to successfully record her students' grades.

### grading

You will be graded on:
- the architecture and elegance of your code
- the cleanliness and documentation of your code
- the number of unit tests that you pass

It is not necessarily expected that you will be able to construct a solution that can pass all the unittests, so focus more on quality if you feel you will have a hard time completing the project.

## python setup

For your convenience, scripts for setting up a Python virtual environment are provided. You must have python installed to use these scripts.

### MacOS and Unix

To setup a Python 2.7 environment on MacOS or linux use the ``setup.sh`` script.

```bash
hostname:username$ sh setup.sh
```

For a Python 3 environment, use the ``setup3.sh`` script. Note that the ``openpyxl`` dependency requires Python >= 3.5.

### Windows

To setup a Python 2.7 environment on Windows use the ``setup.bat`` script.

```batch
C:\Users\username\Documents\roster> setup.bat
```

For a Python 3 environment, use the ``setup3.bat`` script. Note that the ``openpyxl`` dependency requires Python >= 3.5.

## repository setup

In order to share your work with me for review, you will need to setup a private GitHub repository. Here is how that is done:

1. Set up a new private `roster` repository under your personal GitHub account. (These used to be available only to paid users of GitHub but are now available to all users with a limit on collaborators.) Don't forget to make it private, and make sure you don't check "Initialize this repository with a README or include a .gitignore or license file.
1. Add `timothy-thomas` as a collaborator for this new private repository.
1. Clone this repository to your local machine.
1. Navigate to the cloned repository, checkout a new branch to track your changes, add a remote for your private repository, and finally push your new branch to the master branch of your private repository:

```bash
$ git checkout -b private
Switched to a new branch 'private'
$ git remote add private git@github.com:<github-username>/roster.git
$ git push private -u master
Enter passphrase for key '/Users/<user>/.ssh/id_rsa': 
Counting objects: 42, done.
Delta compression using up to 8 threads.
Compressing objects: 100% (36/36), done.
Writing objects: 100% (42/42), 21.90 KiB | 0 bytes/s, done.
Total 42 (delta 17), reused 0 (delta 0)
remote: Resolving deltas: 100% (17/17), done.
To git@github.com:<github-username>/roster.git
 * [new branch]      private -> master
Branch private set up to track remote branch master from private.
```

Now when you commit changes to the `private` branch and push them, they will go to your private repository and be accessible for review by me.
