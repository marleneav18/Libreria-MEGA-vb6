CREATE DATABASE LibreriaMega;

USE LibreriaMega;

CREATE TABLE Generos (
		GeneroID INT PRIMARY KEY IDENTITY,
		Nombre VARCHAR(50) NOT NULL,
		EsFavorito BIT NOT NULL DEFAULT 0
	);

CREATE TABLE Libros (
	LibroID INT PRIMARY KEY IDENTITY,
	Titulo VARCHAR(255) NOT NULL,
	Autor VARCHAR(255) NOT NULL,
	GeneroID INT NOT NULL,
	Calificacion INT NULL,
	Leido BIT NOT NULL DEFAULT 0,
	PorLeer BIT NOT NULL DEFAULT 0,
	Recomendado BIT NOT NULL DEFAULT 0,
	Prestado BIT NOT NULL DEFAULT 0,
	PrestadoA VARCHAR(100) NULL,
	FechaPrestamo DATETIME NULL,
	FOREIGN KEY (GeneroID) REFERENCES Generos(GeneroID)
);

INSERT INTO Libros(Titulo, Autor, GeneroID, Calificacion, Leido, PorLeer, Recomendado, Prestado, PrestadoA, FechaPrestamo)
VALUES
	('El Principito', 'Antoine De Saint', 5, 5, 1, 0, 1, 1, 'Marlene Aguilar', '2025-05-25'),
	('Alicia en el pais de las maravillas', 'Lewis Carroll', 5, 4, 0, 1, 1, 0, NULL, NULL),
	('Dracula', 'Bram Stoker', 1, 5, 1, 0, 1, 0, NULL, NULL);
	
