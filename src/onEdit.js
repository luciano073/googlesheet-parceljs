function onEdit(e) {
  if(e.source.getSheetName() !== 'Formularios') return
  if(e.range.getA1Notation() == 'C4') searchBalance()
  if(e.range.getA1Notation() == 'I4') searchCompraBRL()
  if(e.range.getA1Notation() == 'L4') searchVendaBRL()
  if(e.range.getA1Notation() == 'C21') searchProventos()
}
