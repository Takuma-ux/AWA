from spire.doc import *
from spire.doc.common import *

# Find all the hyperlinks in a document
def FindAllHyperlinks(document):
    hyperlinks = []
    for i in range(document.Sections.Count):
        section = document.Sections.get_Item(i)
        for j in range(section.Body.ChildObjects.Count):
            sec = section.Body.ChildObjects.get_Item(j)
            if sec.DocumentObjectType == DocumentObjectType.Paragraph:
                for k in range((sec if isinstance(sec, Paragraph) else None).ChildObjects.Count):
                    para = (sec if isinstance(sec, Paragraph) else None).ChildObjects.get_Item(k)
                    if para.DocumentObjectType == DocumentObjectType.Field:
                        field = para if isinstance(para, Field) else None
                        if field.Type == FieldType.FieldHyperlink:
                            hyperlinks.append(field)
    return hyperlinks

# Flatten the hyperlink fields
def FlattenHyperlinks(field):
    ownerParaIndex = field.OwnerParagraph.OwnerTextBody.ChildObjects.IndexOf(field.OwnerParagraph)
    fieldIndex = field.OwnerParagraph.ChildObjects.IndexOf(field)
    sepOwnerPara = field.Separator.OwnerParagraph
    sepOwnerParaIndex = field.Separator.OwnerParagraph.OwnerTextBody.ChildObjects.IndexOf(field.Separator.OwnerParagraph)
    sepIndex = field.Separator.OwnerParagraph.ChildObjects.IndexOf(field.Separator)
    endIndex = field.End.OwnerParagraph.ChildObjects.IndexOf(field.End)
    endOwnerParaIndex = field.End.OwnerParagraph.OwnerTextBody.ChildObjects.IndexOf(field.End.OwnerParagraph)

    FormatFieldResultText(field.Separator.OwnerParagraph.OwnerTextBody, sepOwnerParaIndex, endOwnerParaIndex, sepIndex, endIndex)

    field.End.OwnerParagraph.ChildObjects.RemoveAt(endIndex)

    for i in range(sepOwnerParaIndex, ownerParaIndex - 1, -1):
        if i == sepOwnerParaIndex and i == ownerParaIndex:
            for j in range(sepIndex, fieldIndex - 1, -1):
                field.OwnerParagraph.ChildObjects.RemoveAt(j)
        elif i == ownerParaIndex:
            for j in range(field.OwnerParagraph.ChildObjects.Count - 1, fieldIndex - 1, -1):
                field.OwnerParagraph.ChildObjects.RemoveAt(j)
        elif i == sepOwnerParaIndex:
            for j in range(sepIndex, -1, -1):
                sepOwnerPara.ChildObjects.RemoveAt(j)
        else:
            field.OwnerParagraph.OwnerTextBody.ChildObjects.RemoveAt(i)

# Convert fields to text range and clear the text formatting
def FormatFieldResultText(ownerBody, sepOwnerParaIndex, endOwnerParaIndex, sepIndex, endIndex):
    for i in range(sepOwnerParaIndex, endOwnerParaIndex + 1):
        para = ownerBody.ChildObjects[i] if isinstance(ownerBody.ChildObjects[i], Paragraph) else None
        if i == sepOwnerParaIndex and i == endOwnerParaIndex:
            for j in range(sepIndex + 1, endIndex):
               if isinstance(para.ChildObjects[j], TextRange):
                 FormatText(para.ChildObjects[j])
        elif i == sepOwnerParaIndex:
            for j in range(sepIndex + 1, para.ChildObjects.Count):
                if isinstance(para.ChildObjects[j], TextRange):
                  FormatText(para.ChildObjects[j])
        elif i == endOwnerParaIndex:
            for j in range(0, endIndex):
               if isinstance(para.ChildObjects[j], TextRange):
                 FormatText(para.ChildObjects[j])
        else:
            for j, unusedItem in enumerate(para.ChildObjects):
                if isinstance(para.ChildObjects[j], TextRange):
                  FormatText(para.ChildObjects[j])

# Format text
def FormatText(tr):
    tr.CharacterFormat.TextColor = Color.get_Black()
    tr.CharacterFormat.UnderlineStyle = UnderlineStyle.none

# Create a Document object
doc = Document()

# Load a Word file
word_file_path = r'C:\Users\takum\OneDrive\ドキュメント\AWA\takara\output\107\240213【初稿】ビルトインコンロとは_without_toc_final_no_images.docx'
doc.LoadFromFile(word_file_path)

# Get all hyperlinks
hyperlinks = FindAllHyperlinks(doc)

# Flatten all hyperlinks
for i in range(len(hyperlinks) - 1, -1, -1):
    FlattenHyperlinks(hyperlinks[i])
output_file_path = r'C:\Users\takum\OneDrive\ドキュメント\AWA\takara\output\107\remove_hyperlinks.docx'
# Save to a different file
doc.SaveToFile(output_file_path, FileFormat.Docx)
doc.Close()
