# Flash Card Generator :zap::black_joker: 

This Python script can be used to generate printable flashcards as a Microsoft Word document.

The purpose of the script is to demonstrate how MS Word docs can be automatically generated based on an input MS Excel file.

## Dependencies

To run this script you will need ```Python 3``` with the following modules installed:

```bash
pip install xlrd==1.2.0
pip install python-docx==0.8.10
```

## How to use

1. Open ```Input_Data.xlsx```
2. Fill **Column A** with the text that you want to see on each card's **top side**
3. Fill **Column B** with the text that you want to see on each card's **bottom side**
4. When you are done save the file
5. Run ```flash_card_generator.py``` script
6. When the script is done a new file called ```Printable_Flash_Cards.docx``` will appear in the root folder
7. Open the freshly generated .docx file, make sure it is exactly what you wanted
8. Print the document as doublesided
9. Enjoy your brand new flash cards 

### Content of the generated file

The generated output file (Printable_Flash_Cards.docx) has the ollowing parts:
1. Cover page
2. Printable cards
3. A summary table with the content of every flash card.

### Bulk changing the style of the cards

If you want to change the font, size etc. of the cards at once you can do easily by modifying the pre-define Word Style called: "Normal".

### File names

Change the default file names with the help of the below global varables:

```python
INPUT_FILE_NAME     = 'Input_Data.xlsx'
TEMPLATE_FILE_NAME  = 'Template.docx'
OUTPUT_FILE_NAME    = 'Printable_Flash_Cards.docx'
```

## TODO

### New features

- [ ] Make card dimensions modifyable. Currently it can be reduced but every page will consist of 12 cards.
- [ ] Make the number of cards on a single page modifyable.

### Known bugs and errors

TBD