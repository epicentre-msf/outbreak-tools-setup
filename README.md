

<div align="center">
<p align="center">

## *Setup File* for [outbreak tools](https://github.com/epicentre-msf/outbreak-tools) :hammer:

[![Download stable version of setup file](https://github.com/epicentre-msf/outbreak-tools-setup/raw/main/src/imgs/stable_setup.svg)](https://github.com/epicentre-msf/outbreak-tools-setup/raw/main/setup.xlsb)
[![Download development version of setup file](https://github.com/epicentre-msf/outbreak-tools-setup/raw/main/src/imgs/dev_setup.svg)](https://github.com/epicentre-msf/outbreak-tools-setup/raw/dev/src/bin/setup_dev.xlsb)
[![Documentation](https://github.com/epicentre-msf/outbreak-tools/raw/main/src/imgs/docs.svg)](https://github.com/epicentre-msf/outbreak-tools-setup/wiki)

</p>
</div>
The setup file is a configuration `.xlsb` file where you define the overall elements of your linelist. It has five main sheets:

- The *Dictionary* sheet with variable names and other properties
- The *Choice* sheet where you define dropdowns to be used in the linelist
- The *Analysis* sheet where you can configure analysis
- The *Export* sheet for configurations on exports
- The *Translation* sheet where you deal with languages to be used in the linelist.

This repo is the setup file's repo for outbreak tools.
For documentations on how to use the setup file, please refer to the outbreak-tools [documentation](https://epicentre-msf.github.io/outbreak-tools).

`- src` contains source files for the setup

-` setup.xlsb` is the setup file with test data in it. You can import data from another setup file and check a setup for eventual errors.

