# commission-comparer-infynity

## Generating a binary file
NOTE: Make sure you generate the binary in the same **operating system** it is going to be used and ensure the **python version** of the project is the same installed in your OS.

1. Clone project from git
1. Start your development environment
1. Install all dependencies in your environment
1. Generate binary file
    1. Run `pyinstaller name_of_main_file.py --onefile` inside the project directory

       This will create two new directories inside your project `build` and `dist`.

    1. Your cli binary will be inside the `dist` directory

       You can rename the cli file to anything you want.

       Don't forget to make it an executable by running `chmod +x [name_of_file]`.
1. Now try running `python cli.py --help` to have a list of commands.

> DEV NOTES: It will not work because the path to the generated files is pointing to inside the project at the moment.
