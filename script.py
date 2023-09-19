import openpyxl
import win32print

source_file_path = "source.xlsx"
destination_file_path = "destination.xlsx"

def copy_cell_value(source_file_path, destination_file_path, source_cell_address, destination_cell_address):
  """Copies the value of a cell from one XLSX file to another.

  Args:
    source_file_path: The path to the source XLSX file.
    destination_file_path: The path to the destination XLSX file.
    source_cell_address: The address of the cell in the source file to copy.
    destination_cell_address: The address of the cell in the destination file to write to.
  """

  # Open the source and destination XLSX files.
  source_workbook = openpyxl.load_workbook(source_file_path)
  destination_workbook = openpyxl.load_workbook(destination_file_path)

  # Get the worksheets from the source and destination workbooks.
  source_worksheet = source_workbook.worksheets[0]
  destination_worksheet = destination_workbook.worksheets[0]

  # Get the value of the cell in the source file to copy.
  source_cell_value = source_worksheet[source_cell_address].value

  # Write the value of the cell from the source file to the destination file.
  destination_worksheet[destination_cell_address].value = source_cell_value

  # Save the destination workbook.
  destination_workbook.save(destination_file_path)

# Copy the value of cell A1 from the source file to A1 of the destination file.
copy_cell_value(source_file_path, destination_file_path, "A1", "A1")

def print_xl_file(file_path, printer_name):
  """Prints an XL file to the specified printer.

  Args:
    file_path: The path to the XL file to print.
    printer_name: The name of the printer to print to.
  """

  # Get the default printer.
  default_printer = win32print.GetDefaultPrinter()

  # Set the selected printer.
  win32print.SetDefaultPrinter(printer_name)

  # Print the XL file.
  win32print.ShellExecute(0, "printto", file_path, None, None, 0)

  # Set the default printer back to the original printer.
  win32print.SetDefaultPrinter(default_printer)

def select_printer():
  """Prompts the user to select a printer from the list of available printers.

  Returns:
    The name of the selected printer.
  """

  printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL, None, 2)

  print("Please select a printer:")
  for i in range(len(printers)):
    print(f"{i}. {printers[i]['pPrinterName']}")

  printer_index = int(input("Enter the printer index: "))

  return printers[printer_index]['pPrinterName']

# Example usage:

file_path = "my_file.xlsx"

# Select a printer from the list of available printers.
printer_name = select_printer()

# Print the XL file to the selected printer.
print_xl_file(file_path, printer_name)