function getColumnLetter(columnIndex) {
  let columnLetter = "";
  let divisor = columnIndex;
  let modulo;

  while (divisor > 0) {
    modulo = (divisor - 1) % 26;
    columnLetter = String.fromCharCode(65 + modulo) + columnLetter;
    divisor = Math.floor((divisor - modulo) / 26);
  }

  return columnLetter;
}
