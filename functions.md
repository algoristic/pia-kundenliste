# Formeln

## Daten (_Kundenliste_)

### Spalte *Bestellung notwendig*

```
=
WENN(
	I2="";
	"";
	I2+(
		WENNFEHLER(
			WENNNV(
				WENN(
					MIN(
						WENN(
							SVERWEIS(
								B2;
								Stammdaten!A:K;
								7
							)>0;
							SVERWEIS(
								B2;
								Stammdaten!A:K;
								7
							);
							9999
						);
						WENN(
							SVERWEIS(
								B2;
								Stammdaten!A:K;
								11
							)>0;
							SVERWEIS(
								B2;
								Stammdaten!A:K;
								11
							);
							9999
						)
					)=9999;
					30;
					MIN(
						WENN(
							SVERWEIS(
								B2;
								Stammdaten!A:K;
								7
							)>0;
							SVERWEIS(
								B2;
								Stammdaten!A:K;
								7
							);
							9999
						);
						WENN(
							SVERWEIS(
								B2;
								Stammdaten!A:K;
								11
							)>0;
							SVERWEIS(
								B2;
								Stammdaten!A:K;
								11
							);
							9999
						)
					)
				);
				30
			);
			30
		)
	)
)
```
```
=WENN(I2="";"";I2+(WENNFEHLER(WENNNV(WENN(MIN(WENN(SVERWEIS(B2;Stammdaten!A:K;7)>0;SVERWEIS(B2;Stammdaten!A:K;7);9999);WENN(SVERWEIS(B2;Stammdaten!A:K;11)>0;SVERWEIS(B2;Stammdaten!A:K;11);9999))=9999;30;MIN(WENN(SVERWEIS(B2;Stammdaten!A:K;7)>0;SVERWEIS(B2;Stammdaten!A:K;7);9999);WENN(SVERWEIS(B2;Stammdaten!A:K;11)>0;SVERWEIS(B2;Stammdaten!A:K;11);9999)));30);30)))
```

### Spalte *Vorrat für [Tage / Hund]*

```
=
WENNFEHLER(
	RUNDEN(
		(
			(E2*6)*D2
		)/F2;
		0
	);
	""
)
```
```
=WENNFEHLER(RUNDEN(((E2*6)*D2)/F2; 0);"")
```

### Spalte *Vorrat für [Tage / Katze]*

```
=
WENNFEHLER(
	RUNDEN(
		(
			(I2*6)*H2
		)/J2;
		0
	);
	""
)
```
```
=WENNFEHLER(RUNDEN(((I2*6)*H2)/J2; 0);"")
```

## Bedingte Formatierung (_Kundenliste_)

### Spalte *Bestellung notwendig*

`=UND(ISTZAHL(J1);DATEDIF(HEUTE();J1;"D")>14)`: `#007256`

`=UND(ISTZAHL(J1);DATEDIF(HEUTE();J1;"D")<=14;DATEDIF(HEUTE();J1;"D")>7)`: `#F6C700`

`=UND(ISTZAHL(J1);DATEDIF(HEUTE();J1;"D")<=7)`: `#BD1E24`

### Spalte *Erstkontakt*

`=UND(ISTZAHL(G1);DATEDIF(HEUTE();G1;"D")>7)`: `#BD1E24`

`=UND(ISTZAHL(G1);DATEDIF(HEUTE();G1;"D")<=7;DATEDIF(HEUTE();G1;"D")>2)`: `#F6C700`

`=UND(ISTZAHL(G1);DATEDIF(HEUTE();G1;"D")<=2)`: `#007256`

## Bedingte Formatierung (_Geburtstage_)

`=UND(ISTZAHL(A1);DATEDIF(HEUTE();A1;"D")>20)`: `#007256`

`=UND(ISTZAHL(A1);DATEDIF(HEUTE();A1;"D")<=20;DATEDIF(HEUTE();A1;"D")>10)`: `#F6C700`

`=UND(ISTZAHL(A1);DATEDIF(HEUTE();A1;"D")<=10)`: `#BD1E24`
