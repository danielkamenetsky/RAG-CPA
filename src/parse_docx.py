from docx import Document
import os

# This function is used to extract the text and tables from the docx file
def extract_text_and_tables(docx_path):
    # Creating an instance of the Document class and passing the path to the docx file
    doc = Document(docx_path)
    # Creating an empty list to store the elements of the document
    elements = []

    for block in iter_block_items(doc):
        # isinstance is used to check if the block is a string
        if isinstance(block, str):
            # if block checks if the block is not empty
            if block.strip():
                elements.append({"type": "text", "content": block.strip()})
        elif isinstance(block, list):  # table rows
            # this complex variable is used to iterate through the rows of the table
            # and join the cells with a pipe symbol
            table_text = "\n".join([" | ".join(cell.strip() for cell in row) for row in block])
            elements.append({"type": "table", "content": table_text})

    return elements

def iter_block_items(doc):
    """
    Yield paragraphs and tables from the document in order.
    """
    # The imports are inside the function to avoid circular imports.
    # Circular imports are when a module imports itself. 
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    for child in doc.element.body.iterchildren():
        if child.tag.endswith('}p'):
            yield Paragraph(child, doc)
        elif child.tag.endswith('}tbl'):
            table = Table(child, doc)
            table_data = []
            for row in table.rows:
                table_data.append([cell.text for cell in row.cells])
            yield table_data

if __name__ == "__main__":
    # Option 1: Fix the relative path if needed
    docx_path = "data/ifrs16.docx"
    
    # Option 2: Use an absolute path (replace with your actual path)
    # docx_path = "/Users/danielkamenetsky/GithubRepos/rag-cpa/data/ifrs16.docx"
    
    # Option 3: Use os.path to make paths more reliable
    # import os
    # docx_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data", "ifrs16.docx")
    
    # Add a check to verify the file exists before trying to process it
    if not os.path.exists(docx_path):
        print(f"Error: File not found at '{docx_path}'")
        print(f"Current working directory: {os.getcwd()}")
    else:
        elements = extract_text_and_tables(docx_path)
        # Preview the first few elements
        for elem in elements[:5]:
            print(f"[{elem['type']}] {elem['content'][:200]}")
