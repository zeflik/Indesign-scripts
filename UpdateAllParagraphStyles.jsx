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
            var paragraphHyphenation = checkboxControls.add({
              staticLabel: "Hyphenation",
              checkedState: true
            });
          }
          with(dialogRows.add()) {
            staticTexts.add({
              staticLabel: "Words with at Least:"
            });
            var wordsWithAtLeast = integerEditboxes.add({
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
            var afterFirst = integerEditboxes.add({
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
            var beforeLast = integerEditboxes.add({
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
            var hyphenLimit = integerEditboxes.add({
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
            var hyphenationZone = integerEditboxes.add({
              editContents: "0"
            });
          }
          with(dialogRows.add()) {
            var hyphenateCapitalizedWords = checkboxControls.add({
              staticLabel: "Hyphenate Capitalized Words",
              checkedState: false
            });
          }
          with(dialogRows.add()) {
            var lastWordHypenation = checkboxControls.add({
              staticLabel: "Hyphenate Last Word",
              checkedState: false
            });
          }
          with(dialogRows.add()) {
            var hyphenateAcrossColumn = checkboxControls.add({
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
    var paragraph;
    for (i = 1; i < paragraphStyles.length; i++) {
      try {
        style = paragraphStyles[i];
        if (languageDropdown.selectedIndex > 0) style.appliedLanguage = languages[languageDropdown.selectedIndex];
        if (changeToParagraphComposer) style.composer = 'Adobe Paragraph Composer';
        style.hyphenateLastWord = lastWordHypenation.checkedState;
        style.hyphenation = paragraphHyphenation.checkedState;
        style.hyphenateAfterFirst = afterFirst.editValue;
        style.hyphenateBeforeLast = beforeLast.editValue;
        style.hyphenateCapitalizedWords = hyphenateCapitalizedWords.checkedState;
        style.hyphenateLadderLimit = hyphenLimit.editValue;
        style.hyphenateLastWord = lastWordHypenation.checkedState;
        style.hyphenateWordsLongerThan = wordsWithAtLeast.editValue;
        style.hyphenationZone = hyphenationZone.editValue;
        style.hyphenateAcrossColumns = hyphenateAcrossColumn.checkedState;
      } catch (error) {
        alert(error, APP_NAME);
        return;
      }
    }
  }
}
