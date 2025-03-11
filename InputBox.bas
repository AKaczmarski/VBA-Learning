Attribute VB_Name = "Module1"
Sub formularz()

x = InputBox("Podaj imie", "Formularz osobowy", "imie", 1000, 1000, "pomoc.hlp", 10)
Range("A1") = x

Range("A2") = InputBox("Podaj nazwisko", "Formularz osobowy", "nazwisko", 1100, 1100, "pomoc.hlp", 10)

End Sub
