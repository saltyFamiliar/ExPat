import copy
from docx import Document

#recursively looks through a Word doc and replaces targets with data from patient_dict
def find_and_replace(doc, targets, patient_dict):
		for table in doc.tables:
			for row in table.rows:
				for cell in row.cells:
					for paragraph in cell.paragraphs:
						for target in targets:
							if str(target) in paragraph.text:
								paragraph.text = paragraph.text.replace(f"#{target}", patient_dict[target])
					find_and_replace(cell, targets, patient_dict)


#wrapper for find_and_replace that handles file creation
def create_file(patient_dict, targets, document, new_document_name):
	newDocument = copy.deepcopy(document)
	find_and_replace(newDocument, targets, patient_dict)
	newDocument.save(new_document_name)
	return 1
