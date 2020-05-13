# pyenv for Windows

[pyenv][1] is a great tool. I ported it to Windows. Some commands aren't implemented, but it's good enough for basic use.

For existing python users, we support installation via pip: [follow instructions](#installation)  
Contributors and Interested people can join us @ [Slack](https://join.slack.com/t/pyenvforwindows/shared_invite/enQtODQ0NTQwMjEzOTA4LWE1YjMwMGI3YTAwYjg0MGQzNjg1ZjJiYTMzMTJkNWM4NTY5OTY4YmU3ZjhjMTc3MGJhMDY1MmJkMjFmY2ViNTI)

>  Link to Official [2019 Python Developers Survey](https://surveys.jetbrains.com/s3/c16-python-developers-survey-2019)

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![GitHub issues open](https://img.shields.io/github/issues/pyenv-win/pyenv-win.svg?)](https://github.com/pyenv-win/pyenv-win/issues)
[![Downloads](https://pepy.tech/badge/pyenv-win)](https://pepy.tech/project/pyenv-win)

- [Introduction](#introduction)
- [pyenv](#pyenv)
- [pyenv-win commands](#pyenv-win-commands)
- [Installation](#installation)
   - [Get pyenv-win](#get-pyenv-win)
   - [Finish the installation](#finish-the-installation)
- [Usage](#usage)
- [How to get updates](#how-to-get-updates)
- [FAQ](#faq)
- [How to contribute](#how-to-contribute)
- [Bug Tracker and Support](#bug-tracker-and-support)
- [License and Copyright](#license-and-copyright)
- [Author and Thanks](#author-and-thanks)

## Introduction

[pyenv][1] for python is a great tool but, like [rbenv][2] for ruby developers, it doesn't support Windows directly. After a bit of research and feedback from python developers, I discovered they wanted a similiar feature for Windows systems.

I got inspired from the pyenv [issue][4] for Windows support. Personally, I use Mac and Linux with beautiful [pyenv][1], but some companies still use Windows for development. This library is to help Windows users manage multiple python versions.

I found a similar system for [rbenv-win][3] for ruby developers. This project was forked from [rbenv-win][3] and modified for [pyenv][1]. Some commands aren't implemented, but it's good enough for basic use.

## pyenv

[pyenv][1] is a simple python version management tool. It lets you easily switch between multiple versions of Python. It's simple, unobtrusive, and follows the UNIX tradition of single-purpose tools that do one thing well.

## pyenv-win commands

```yml
   commands     List all available pyenv commands
   duplicate    Creates a duplicate python environment
   local        Set or show the local application-specific Python versions
   global       Set or show the global Python versions
   shell        Set or show the shell-specific Python versions
   install      Install one or more Python versions
   uninstall    Uninstall one or more Python versions
   update       Update the cached version DB
   rehash       Rehash pyenv shims (run this after installing libraries with pip)
   version      Show the current Python version(s) and their origin
   version-name Show the current Python version(s)
   versions     List all Python versions available to pyenv
   exec         Runs an executable by first preparing PATH so that the 
                  selected Python version executes first
   which        Display the full path to an executable
   whence       List all Python versions that contain the given executable
```

## Installation

### Get pyenv-win

Get pyenv-win via one of the following methods:

- **With pip** (to support existing python users)
   - Powershell or Git Bash: `pip install pyenv-win --target $HOME\.pyenv`
   - cmd.exe: `pip install pyenv-win --target %USERPROFILE%\.pyenv`
- **With zip file**
   1. Download link: [pyenv-win](https://github.com/pyenv-win/pyenv-win/archive/master.zip)
   2. Extract to `%USERPROFILE%\.pyenv\pyenv-win`
- **With Git**
   - Powershell or Git Bash: `git clone https://github.com/pyenv-win/pyenv-win.git ~/.pyenv`
   - cmd.exe: `git clone https://github.com/pyenv-win/pyenv-win.git %USERPROFILE%\.pyenv`
- **With [Chocolatey](https://chocolatey.org/packages/pyenv-win)**
   - `choco install pyenv-win` (this also installs all the environment variables)

### Finish the installation

   1. If you installed using Chocolatey, you can skip to step 3.
      Otherwise, add a new *User* variable in ENVIRONMENT with name:  
      `PYENV` value: `C:\Users\[username]\.pyenv\pyenv-win`
   2. Now add the following paths to your *User* ENVIRONMENT PATH variable in order to access the pyenv command (don't forget to separate with semicolons, and use this exact order for the shim folders):
      - `%PYENV%\bin`
      - `%PYENV%\win-shims`
      - `%PYENV%\nix-shims`
      
      This can be done through GUI or command line:
      - __This PC → Properties → Advanced system settings → Advanced → Environment Variables... → PATH__
      - _Be careful! People who use Windows newer than May 2019 Update must put these items above `%USERPROFILE%\AppData\Local\Microsoft\WindowsApps`; See [this article](https://devblogs.microsoft.com/python/python-in-the-windows-10-may-2019-update/)._
      - Powershell:
        - `[System.Environment]::SetEnvironmentVariable('PYENV', "$env:HOME\.pyenv\pyenv-win", [System.EnvironmentVariableTarget]::User)`
        - `$UserPath = [System.Environment]::GetEnvironmentVariable('PATH', [System.EnvironmentVariableTarget]::User).Trim(";")`
        - `[System.Environment]::SetEnvironmentVariable('PATH', [String]::Concat($UserPath, ";%PYENV%\bin;%PYENV%\win-shims;%PYENV%\nix-shims").Trim(";"), [System.EnvironmentVariableTarget]::User)`
   
   3. Verify the installation was successful by opening a new terminal and running `pyenv --version`
   4. Now run the `pyenv rehash` from home directory
      - You should see the [current pyenv version](https://github.com/pyenv-win/pyenv-win/blob/master/setup.py). If you are getting an error, go through the steps again. Still facing the issue? [Open a ticket](https://github.com/pyenv-win/pyenv-win/issues).
   5. Run `pyenv` to see list of commands it supports. [More info...](#usage)

   Installation is done. Hurray!

## Usage

- Update the list of discoverable Python versions: `pyenv update`
- To view a list of python versions supported by pyenv windows: `pyenv install -l`
- To install a python version:  `pyenv install 3.5.2`
   - _Note: An install wizard may pop up for some non-silent installs. You'll need to click through the wizard during installation. There's no need to change any options in it._
   - You can also install multiple versions in one command too: `pyenv install 2.4.3 3.6.8`
- To set a python version as the global version: `pyenv global 3.5.2`
   - This is the version of python that will be used by default if a local version (see below) isn't set.
   - _Note: The version must first be installed._
- To set a python version as the local version: `pyenv local 3.5.2`.
   - The version given will be used whenever `python` is called from within this folder. This is different than a virtual env, which needs to be explicitly activated.
   - _Note: The version must first be installed._
- After (un)installing any libraries using pip or modifying the files in a version's folder, you must run `pyenv rehash` to update pyenv with new shims for the python and libraries' executables.
   - _Note: This must be run outside of the `.pyenv` folder._
- To uninstall a python version: `pyenv uninstall 3.5.2`
- To view which python you are using and its path: `pyenv version`
- To view all the python versions installed on this system: `pyenv versions`

## How to get updates

- If installed via pip
   - Add pyenv-win installed path to `easy_install.pth` file which is located in site-package. Now pyenv-win is recognised by pip
   - Get updates via pip `pip install --upgrade pyenv-win`
- If installed via Git
   - Go to the `%USERPROFILE%\.pyenv\pyenv-win` (which is your installed path) and run `git pull`
- If installed via zip
   - Download the latest zip and extract it
   - Go to `%USERPROFILE%\.pyenv\pyenv-win` and replace the folders `libexec` and `bin` with the new ones you just downloaded

## FAQ

- **Question:** Does pyenv for windows support python2?
   - **Answer:** Yes, We support python2 from version 2.4+ until python.org officially removes it.
   - Versions below 2.4 use outdated Wise installers and have issues installing multiple patch versions, unlike Windows MSI and the new Python3 installers that support "extraction" installations.

- **Question:** Does pyenv for windows support python3?
   - **Answer:** Yes, we support python3 from version 3.0. We support it from 3.0 until python.org officially removes it.

- **Question:** I am getting the issue `batch file cannot be found.` while installing python, what should I do?
   - **Answer:** You can ignore it. It's calling `pyenv rehash` command before creating the bat file in few devices.

- **Question:** System is stuck while uninstalling the python version, what to do?
   - **Answer:** Navigate to the location you installed pyenv, opening its 'versions' folder (usually `%USERPROFILE%\.pyenv\pyenv-win\versions`), and delete the folder of the version you want removed.

- **Question:** I installed pyenv-win using pip. How can I uninstall it?
   - **Answer:** Follow the pip instructions in [How to get updates](#how-to-get-updates) and then run `pip uninstall pyenv-win`

- **Question:** pyenv-win not recognised, but I have set the ENV PATH?
   - **Answer:** According to Windows when adding the path under the User or System variables, for the User variable you need to logout and login again to reflect the changes. For System variable it's not required.

## Change Log

### New in 1.4.0
- Added basic support for using multiple instances of python after a rehash via "version specific" shims (for compatability with libraries like tox that expect them). For example, `pyenv global 3.6.8 3.7.7 3.8.2 3.5.4 2.7.17 2.6.6` would make the following Python commands available and execute these specific versions when called:
   - python    (3.6.8)
   - python2   (2.7.17)
   - python2.6 (2.6.6)
   - python2.7 (2.7.17)
   - python3   (3.6.8)
   - python3.5 (3.5.4)
   - python3.6 (3.6.8)
   - python3.7 (3.7.7)
   - python3.8 (3.8.2)
   - **\*Note\*: The shims are setup by order of first occurring version in the `pyenv global/local/shell` command.**
- Split the `pyenv-win\shims` folder into their respective sub-shell folders (`pyenv-win\nix-shims` and `pyenv-win\win-shims`) to support calling the new "version specific" shims for python directly in Command Prompt.
   - Otherwise you'd have to append `.bat` to each call for a specific python version when using Command Prompt. (`python3.7.bat -m timeit` instead of `python3.7 -m timeit`)
   - **\*WARNING\*: This change is backwards incompatible with v1.3.0 and will require you to redo your environment variable setup (you should delete `pyenv-win\shims` and update your %PATH% variable).**
- `pyenv rehash` is automatically performed when calling `pyenv global/local` now, you don't have to use `pyenv global [version] && pyenv rehash` anymore.

### New in 1.3.0
- Version naming conventions have now changed from using 64-bit suffixes when specifying a version to (un)install. Now all you need to use is the version number to install your platform's specifc bit version.
   - **\*WARNING\*: This change is backwards incompatible with v1.2.5 so please uninstall all versions of python prior to upgrading pyenv to this version.**
   - Ex. `pyenv install 2.7.17` will install as 64-bit on x64 and 32-bit on x86. (64-bit can still use `2.7.17-win32` to install the 32-bit version)
   - `pyenv global/local/shell` also now recognize your platform and select the appropirate bit version. (64-bit users will need to specify `[version]-win32` to use the 32-bit versions now)
- Added support for true unobtrusive, local installs.
  - **\*WARNING\*: This change is backwards incompatible with v1.2.5 so please uninstall all versions of python prior to upgrading pyenv to this version.**
  - No install/uninstall records are written to the registry or Start Menu anymore (no "Programs and Features" records).
  - When installing a patch version of python (ex. 3.6.1) installing another patch version (ex. 3.6.2) won't reuse the same folder and overwrite the previously installed minor version. They're now kept separate.
  - Uninstalls are now a simple folder deletion. (Can be done manually by the user safely now or `pyenv uninstall`)
- Added support for (un)installing multiple versions of python in a single command or all DB versions via the `-a/--all` switch.
   - When using `--all` on x64 computers you can use `--32only` or `--64only` to install only 32-bit or only 64-bit version s of python. (Does nothing on 32-bit computers, and better filters may be in the works later on)
- `pyenv global/rehash` is called automatically after (un)installing a new Python version. (last version specified, if installing multiple)
- Pyenv now uses a cached DB of versions scraped straight from the Python mirror site and can be updated manually by a user using `pyenv update`. Users no longer have to wait for pyenv's source repo to be updated to use a new version of Python when it releases, and can also use the new alpha/beta python releases.
- `pyenv install` now has a `-c/--clear` to empty cached installers in the `%PYENV%\install_cache` folder.
- `pyenv rehash` now acknowledges %PATHEXT% (plus PY and PYW) when creating shims instead of just for exe, bat, cmd and py files so more executables are available from `\Scripts` and libraries installed using pip.
- Shims created using `pyenv rehash` no longer call `pyenv exec`, but instead call python directly to prevent issues with other programs executing the shims.
- Shims now use cp1250 as the default code page since Python2 will [never actually support cp65001](https://bugs.python.org/issue6058#msg120712). cp1250 has better support for upper ANSI characters (ex. "Pokémon"), but still isn't full UTF-8 compatible.
- **Note: Support for Python versions below 2.4 have been dropped since their installers don't install "cleanly" like versions from 2.4 onward and they're predominently out of use/support in most envrionments now.**

## How to contribute

- Fork the project & clone locally.
- Create an upstream remote and sync your local copy before you branch.
- Branch for each separate piece of work. It's a good practise to write test cases.
- Do the work, write good commit messages, and read the CONTRIBUTING file if there is one.
- Test the changes by running `tests\test_install.bat` and `tests\test_uninstall.bat`
- Push to your origin repository.
- Create a new Pull Request in GitHub.

## Bug Tracker and Support

- Please report any suggestions, bug reports, or annoyances with pyenv-win through the [GitHub bug tracker](https://github.com/pyenv-win/pyenv-win/issues).

## License and Copyright

- pyenv-win is licensed under [MIT](http://opensource.org/licenses/mit-license.php) *2019*

   [![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

## Author and Thanks

pyenv-win was developed by [Kiran Kumar Kotari](https://github.com/kirankotari) and [Contributors](https://github.com/pyenv-win/pyenv-win/graphs/contributors)

[1]: https://github.com/pyenv/pyenv
[2]: https://github.com/rbenv/rbenv
[3]: https://github.com/nak1114/rbenv-win
[4]: https://github.com/pyenv/pyenv/issues/62
