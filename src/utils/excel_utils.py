def get_cell_value(sheet, string):
  value = sheet[string].value
  if (value is None):
    return ''
  return value
