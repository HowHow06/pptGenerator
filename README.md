# pptGenerator

PPT generator using Microsoft Excel VBA, to produce Praise and Worship PPT



# Table of Contents

  * [How to generate PPT with Chinese Lyric and Pinyin(or English Lyric)?](#how-to-generate-ppt-with-chinese-lyric-and-pinyin-or-english-lyric)
    + [Sample output:](#sample-output)
    + [Steps:](#steps)
  * [How to generate PPT with Chinese Lyric only? ( Generating Two Text Boxes)](#how-to-generate-ppt-with-chinese-lyric-only--generating-two-text-boxes)
    + [Sample output:](#sample-output-1)
    + [Steps:](#steps-1)
  * [How to generate PPT with Chinese Lyric only? ( Generating OneText Box, Customizabe Row Count)](#how-to-generate-ppt-with-chinese-lyric-only--generating-onetext-box-customizabe-row-count)
    + [Sample Output:](#sample-output-2)
    + [Steps:](#steps-2)

- [Documentation](#documentation)
  * [Chinese Lyric Only Option](#chinese-lyric-only-option)
  * [Use One Text Box Option](#use-one-text-box-option)
  * [File Paths](#file-paths)
    + [Relative Path](#relative-path)
    + [Absolute Path](#absolute-path)
  * [Main and Sub](#main-and-sub)
  * [Font](#font)
  * [Color](#color)
  * [Border Weight](#border-weight)
  * [Font Spacing](#font-spacing)
  * [Line Space](#line-space)
  * [Positioning](#positioning)
  * [Preview](#preview)
  * [Preset](#preset)

# Manual

## How to generate PPT with Chinese Lyric and Pinyin(or English Lyric)?

### Sample output:

![image-20220206175415130](https://raw.githubusercontent.com/HowHow06/Ho2TyporaImages/main/img/image-20220206175415130.png)

### Steps:

1. Select the output PPT file and the background image that will be used. You may manually type in the file path or click the "Select" button to choose the right file.

   By default, the PPT path will be "output.pptx".

   ![image-20220206174953205](https://raw.githubusercontent.com/HowHow06/Ho2TyporaImages/main/img/image-20220206174953205.png)

2. Configure the settings at Sheet 2 (Settings), such as font size, shadow, position etc. Please refer to documentation on the configurations.

   ![image-20220206175232443](https://raw.githubusercontent.com/HowHow06/Ho2TyporaImages/main/img/image-20220206175232443.png)

3. Insert the chinese lyric and pin yin into the text boxes, make sure the lyric is organised nicely line by line:
   ![image-20220206184059420](https://raw.githubusercontent.com/HowHow06/Ho2TyporaImages/main/img/image-20220206184059420.png)

4. Make sure "Chinese Lyric Only" checkbox is not checked
   <img src="https://raw.githubusercontent.com/HowHow06/Ho2TyporaImages/main/img/image-20220206184142128.png" alt="image-20220206184142128"  />

5. Make sure the chinese lyric and pin yin have the same number of rows. For example, if there are 5 lines of chinese lyric, make sure there are also 5 lines of pin yin (including empty line)
   ![image-20220206184417482](https://raw.githubusercontent.com/HowHow06/Ho2TyporaImages/main/img/image-20220206184417482.png)

6. Click generate button.

7. The PPT will be generated and you may Save it as a new file.
   <img src="https://raw.githubusercontent.com/HowHow06/Ho2TyporaImages/main/img/image-20220206184515110.png" alt="image-20220206184515110"  />



## How to generate PPT with Chinese Lyric only? ( Generating Two Text Boxes)

### Sample output:

![image-20220206184713849](https://raw.githubusercontent.com/HowHow06/Ho2TyporaImages/main/img/image-20220206184713849.png)

### Steps:

Follow the Step 1-3 mentioned in the section above: How to generate PPT with Chinese Lyric and Pinyin(or English Lyric)?

4. Check the "Chinese Lyric Only" checkbox but leave the "Use One Text Box" unchecked.
   ![image-20220206184924407](https://raw.githubusercontent.com/HowHow06/Ho2TyporaImages/main/img/image-20220206184924407.png)
5. The input in Pin Yin text box will be ignored. You may leave it blank.
   ![image-20220206185200940](https://raw.githubusercontent.com/HowHow06/Ho2TyporaImages/main/img/image-20220206185200940.png)

Follow the step 6-7 mentioned in the section above: How to generate PPT with Chinese Lyric and Pinyin(or English Lyric)?



## How to generate PPT with Chinese Lyric only? ( Generating OneText Box, Customizabe Row Count)

### Sample Output:

![image-20220206185356802](https://raw.githubusercontent.com/HowHow06/Ho2TyporaImages/main/img/image-20220206185356802.png)

### Steps:

Follow the Step 1-3 mentioned in the section above: How to generate PPT with Chinese Lyric and Pinyin(or English Lyric)?

4. Check the "Chinese Lyric Only" and "Use One Text Box" checkbox. Specify the number of line per PPT Slide. (In this example is 3 lines). 
   ![image-20220206185632006](https://raw.githubusercontent.com/HowHow06/Ho2TyporaImages/main/img/image-20220206185632006.png)
5. The input in Pin Yin text box will be ignored. You may leave it blank.

Follow the step 6-7 mentioned in the section above: How to generate PPT with Chinese Lyric and Pinyin(or English Lyric)?





# Documentation

## Chinese Lyric Only Option

![image-20220206185953461](https://raw.githubusercontent.com/HowHow06/Ho2TyporaImages/main/img/image-20220206185953461.png)

- This PPT Generator was meant to generate PPT slides with Chinese Lyric and Pinyin
- Therefore, to generate PPT slides with Chinese Lyric Only, check the "Chinese Lyric Only" checkbox

## Use One Text Box Option

- by default there will only be 2 lines in each slide. Each line in the slide is placed in seperated textboxes and the position of each textbox can be adjusted in setting (as shown in the diagram below)

  ![image-20220206190254068](https://raw.githubusercontent.com/HowHow06/Ho2TyporaImages/main/img/image-20220206190254068.png)

- Therefore, in order to increase the number of lines per slide, you must use "One Text Box" mode. 

- Once the "Use One Text Box" checkbox is checked, you will need to specify the number of line for each slide
  ![image-20220206190916345](https://raw.githubusercontent.com/HowHow06/Ho2TyporaImages/main/img/image-20220206190916345.png)

- In the sample shown below is one text box with 3 lines of lyric

  ![image-20220206191058234](https://raw.githubusercontent.com/HowHow06/Ho2TyporaImages/main/img/image-20220206191058234.png)



## File Paths

- There are 2 main file paths in this program:
  - PPT Path: The output PPT file
  - Background Image Path: The background image that will be used in output PPT
- Both PPT path and background image path can be absolute or relative
- When the message box below popped up, you might want to make sure your path is correct or click the "Select Output" or "Select Background" button to re-insert the correct file path:
  ![image-20220206174524167](https://raw.githubusercontent.com/HowHow06/Ho2TyporaImages/main/img/image-20220206174524167.png)



### Relative Path

- Example of relative path:
  <img src="https://raw.githubusercontent.com/HowHow06/Ho2TyporaImages/main/img/image-20220206174101117.png" alt="image-20220206191650866"  />

### Absolute Path

- Example of absolute path:
  ![image-20220206174154691](https://raw.githubusercontent.com/HowHow06/Ho2TyporaImages/main/img/image-20220206174154691.png)

## Main and Sub

- Format settings can be done toward Main section and Sub section of the slide at Sheet 2 (Settings)
- ![image-20220206191650866](https://raw.githubusercontent.com/HowHow06/Ho2TyporaImages/main/img/image-image-20220206191650866.png)
- ![image-20220206191658168](https://raw.githubusercontent.com/HowHow06/Ho2TyporaImages/main/img/image-20220206191658168.png)
- ![image-20220206191547109](https://raw.githubusercontent.com/HowHow06/Ho2TyporaImages/main/img/image-20220206191547109.png)

## Font

- The font name used must be exacty the same as the displayed font name in Microsoft Powerpoint
- Example:
  - ![image-20220206191833715](https://raw.githubusercontent.com/HowHow06/Ho2TyporaImages/main/img/image-20220206191833715.png)
  - ![image-20220206191917229](https://raw.githubusercontent.com/HowHow06/Ho2TyporaImages/main/img/image-20220206191917229.png)

## Color

- you may manually type in the RGB value or click the "Pick a Color" button to pick a color



## Border Weight

- ![image-20220206192648897](https://raw.githubusercontent.com/HowHow06/Ho2TyporaImages/main/img/image-20220206192648897.png)
- The thickness of the font outline. (Only visible when "Border" option is checked)
- You may refer to the diagram below
  ![image-20220206192204173](https://raw.githubusercontent.com/HowHow06/Ho2TyporaImages/main/img/image-20220206192204173.png)
- Default value: 1.5 



## Font Spacing

- ![image-20220206192638328](https://raw.githubusercontent.com/HowHow06/Ho2TyporaImages/main/img/image-20220206192638328.png)
- The character spacing, you may adjust accordingly
  ![image-20220206192308539](https://raw.githubusercontent.com/HowHow06/Ho2TyporaImages/main/img/image-20220206192308539.png)
  - Very tight = -3
  - Tight = -1.5
  - Normal = 0
  - Loose = 3
  - Very Loose = 6

## Line Space

- ![image-20220206192851657](https://raw.githubusercontent.com/HowHow06/Ho2TyporaImages/main/img/image-20220206192851657.png)
- The three settings are corresponding to "Spacing Before", "Spacing After", "Line Spacing Multiple At" as shown in the diagram below:
  ![image-20220206192810024](https://raw.githubusercontent.com/HowHow06/Ho2TyporaImages/main/img/image-20220206192810024.png)



## Positioning

- ![image-20220206194131228](https://raw.githubusercontent.com/HowHow06/Ho2TyporaImages/main/img/image-20220206194131228.png)
- Upper and Lower section of Main and Sub can be changed
  - ![image-20220206193207695](https://raw.githubusercontent.com/HowHow06/Ho2TyporaImages/main/img/image-20220206193207695.png)
- ![image-20220206193838259](https://raw.githubusercontent.com/HowHow06/Ho2TyporaImages/main/img/image-20220206193838259.png)
- By Default:
  - Left = 0
  - Top = 70++
  - Height = 90
  - Width = 960 (full width)
- Normally you will only need to adjust the "Top" value, the textbox will move downward when the value is larger



## Preview

- <img src="https://raw.githubusercontent.com/HowHow06/Ho2TyporaImages/main/img/image-20220206194217841.png" alt="image-20220206194217841"  />
- You may preview your setting by inputting the lyric and click preview button



## Preset

- ![image-20220206194317599](https://raw.githubusercontent.com/HowHow06/Ho2TyporaImages/main/img/image-20220206194317599.png)
- After you have completed configuring, you may save the setting to preset (preset 1-6), so that it can be easily loaded 
- You may label the preset in the text box given