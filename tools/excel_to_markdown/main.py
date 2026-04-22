import openpyxl
from typing import Any, Dict

class ExcelToMarkdownTool:
    def _invoke(self, excel_file: Any) -> str:
        """
        Converts the provided Excel file to a Markdown table.
        
        Args:
            excel_file: The uploaded Excel file object.
            
        Returns:
            A string containing the Markdown representation of the first sheet.
        """
        try:
            # Load the workbook
            # excel_file is expected to be a file-like object or a path
            wb = openpyxl.load_workbook(excel_file, data_only=True)
            sheet = wb.active

            rows = list(sheet.rows)
            if not rows:
                return "The Excel file is empty."

            markdown_rows = []
            
            for i, row in enumerate(rows):
                row_data = []
                for cell in row:
                    value = cell.value if cell.value is not None else ""
                    
                    # Check for comments
                    comment = None
                    if cell.comment:
                        comment = cell.comment.text
                    
                    # Format the content: Value (Comment: ...)
                    if comment:
                        cell_content = f"{value} (Comment: {comment})"
                    else:
                        cell_content = str(value)
                    
                    # Escape pipe character to avoid breaking Markdown table
                    cell_content = cell_content.replace("|", "\\|")
                    row_data.append(cell_content)
                
                markdown_rows.append(row_data)

            if not markdown_rows:
                return "No data found in the sheet."

            # Create Markdown table
            # Header
            header = markdown_rows[0]
            markdown_table = "| " + " | ".join(header) + " |\n"
            
            # Separator
            separator = "| " + " | ".join(["---"] * len(header)) + " |\n"
            markdown_table += separator
            
            # Body
            for row in markdown_rows[1:]:
                # Ensure the row has the same number of columns as the header
                if len(row) < len(header):
                    row.extend([""] * (len(header) - len(row)))
                elif len(row) > len(header):
                    row = row[:len(header)]
                
                markdown_table += "| " + " | ".join(row) + " |\n"

            return markdown_table

        except Exception as e:
            return f"Error processing Excel file: {str(e)}"

def main(args: Dict[str, Any]) -> str:
    tool = ExcelToMarkdownTool()
    return tool._invoke(args.get("excel_file"))
