# txt-to-pptx
text file to pptx file converter

This project converts texts into powerpoint slides, with the style of given template pptx file.

This is for Propresenter 7 slide makers to convert multilingual lyrics text to Propresenter slides.

Propresenter 7 is able to import pptx files with the option of 'get text and graphic objects as propresenter slide elements', only in Mac version. (Windows version does not support this option)


## Usage
- Install python and pip
  - For Mac
    - install Homebrew first. https://brew.sh
    - Install python and pip
```
brew install python pip
```
- Install python-pptx
```
pip install python-pptx
```
- Make a 1-page template powerpoint file
- Make txt file of lyrics, separated by blank line for each slide.
- run txt-to-pptx.py
```
python txt-to-pptx.py -t <template.pptx> -i <input.txt> -o <output.pptx>
```
- Open Propresenter 7 in Mac,
  - click 'File-Import-Powerpoint',
  - click 'Option', select 'Import text and graphic object as ProPresenter slide elements',
  - and select converted pptx file.

## Screenshots
![ProPresenter 7 screenshot](./propresenter_screenshot.png?raw=true)
