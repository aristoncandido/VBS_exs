Dim nota1
Dim nota2
Dim titulo

titulo = "Calculador de Media"

nota1 = InputBox("Digite um valor:", titulo)
nota2 = InputBox("Digite mais um valor:", titulo)

Dim resultado
resultado = CDbl(nota1) + CDbl(nota2)

Dim media
media = resultado / 2

MsgBox "A soma dos valores e: " & resultado, vbInformation, titulo
MsgBox "Porem, a media das duas notas e: " & media, vbInformation, titulo
