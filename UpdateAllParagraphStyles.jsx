/*  Created by Józef Niemiec / 1.10.2018
    email: kontakt@jozefniemiec.pl */

const APP_NAME = "Update all Paragraph Styles";

app.doScript(
  main,
  ScriptLanguage.JAVASCRIPT, [],
  UndoModes.ENTIRE_SCRIPT,
  APP_NAME
);

function main() {
  var paragraphStyles = app.activeDocument.allParagraphStyles;
  var languages = app.languagesWithVendors.everyItem().name.sort();
  var myDialog = app.dialogs.add({
    name: APP_NAME,
    canCancel: true,
  });

  with(myDialog) {
    with(dialogColumns.add()) {
      with(dialogRows.add()) {
        staticTexts.add({
          staticLabel: "Set Language"
        });
      }
      var languageDropdown = dropdowns.add({
        stringList: languages
      });
      var changeToParagraphComposer = checkboxControls.add({
        staticLabel: "Set Adobe Paragraph Composer",
        checkedState: false
      });
      with(borderPanels.add()) {
        with(dialogColumns.add()) {
          with(dialogRows.add()) {
            var hyphenationCheckBox = checkboxControls.add({
              staticLabel: "Hyphenation",
              checkedState: true
            });
          }
          with(dialogRows.add()) {
            staticTexts.add({
              staticLabel: "Words with at Least:"
            });
            var wordsWithAtLeastEditText = integerEditboxes.add({
              editContents: "5"
            });
            staticTexts.add({
              staticLabel: "letters",
            });
          }
          with(dialogRows.add()) {
            staticTexts.add({
              staticLabel: "After First:"
            });
            var hyphenateAfterFirstEditText = integerEditboxes.add({
              editContents: "3"
            });
            staticTexts.add({
              staticLabel: "letters"
            });
          }
          with(dialogRows.add()) {
            staticTexts.add({
              staticLabel: "Before Last:"
            });
            var hyphenateBeforeLastEditText = integerEditboxes.add({
              editContents: "3"
            });
            staticTexts.add({
              staticLabel: "letters"
            });
          }
          with(dialogRows.add()) {
            staticTexts.add({
              staticLabel: "Hyphen Limit"
            });
            var hyphenLimitEditText = integerEditboxes.add({
              editContents: "2"
            });
            staticTexts.add({
              staticLabel: "hyphens"
            });
          }
          with(dialogRows.add()) {
            staticTexts.add({
              staticLabel: "Hyphenation Zone"
            });
            var hyphenationZoneEditText = integerEditboxes.add({
              editContents: "0"
            });
          }
          with(dialogRows.add()) {
            var hyphenateCapitalizedCheckBox = checkboxControls.add({
              staticLabel: "Hyphenate Capitalized Words",
              checkedState: false
            });
          }
          with(dialogRows.add()) {
            var hyphenateLastWordCheckBox = checkboxControls.add({
              staticLabel: "Hyphenate Last Word",
              checkedState: false
            });
          }
          with(dialogRows.add()) {
            var hyphenateAcrossColumnCheckbox = checkboxControls.add({
              staticLabel: "Hyphenate Across Column",
              checkedState: false
            });
          }
        }
      }
    }
  }
  var isDialogConfirmed = myDialog.show();
  if (isDialogConfirmed) {
    var paragraphStyle;
    for (i = 1; i < paragraphStyles.length; i++) {
      try {
        paragraphStyle = paragraphStyles[i];
        with(paragraphStyle) {
          if (languageDropdown.selectedIndex > 0) appliedLanguage = languages[languageDropdown.selectedIndex];
          if (changeToParagraphComposer) composer = 'Adobe Paragraph Composer';
          hyphenateLastWord = hyphenateLastWordCheckBox.checkedState;
          hyphenation = hyphenationCheckBox.checkedState;
          hyphenateAfterFirst = hyphenateAfterFirstEditText.editValue;
          hyphenateBeforeLast = hyphenateBeforeLastEditText.editValue;
          hyphenateCapitalizedWords = hyphenateCapitalizedCheckBox.checkedState;
          hyphenateLadderLimit = hyphenLimitEditText.editValue;
          hyphenateWordsLongerThan = wordsWithAtLeastEditText.editValue;
          hyphenationZone = hyphenationZoneEditText.editValue;
          hyphenateAcrossColumns = hyphenateAcrossColumnCheckbox.checkedState;
        }
      } catch (error) {
        alert(error, APP_NAME);
        return;
      }
    }
  }
}
