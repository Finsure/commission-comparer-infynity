# commission-comparer-infynity

## Generating a binary file
> Make sure you generate the binary in the same operating system it is going to be used
> and ensure the python version of the project is the same installed in your OS.

1. Clone project from git
1. Start your development environment
1. Install all dependencies in your environment
1. Generate binary file
    1. Run `pyinstaller name_of_main_file.py --onefile` inside the project directory
    > You will notice there will be created a few new directories inside your project `dist` and `build`
    1. Our cli binary will be inside the `dist` directory
    > You can rename the cli file to anything you want.
    > Don't forget to make it an executable before running `chmod +x [name_of_file]`
1. Now try running `python cli.py`

> DEV NOTES: It may not work because the path to the generated files is pointing to in side the project.
