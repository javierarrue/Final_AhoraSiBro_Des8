create database finalDes5
use finalDes5

create table equipos (
	id_equipo int IDENTITY(1,1), 
	corregimiento varchar(50), 
	PRIMARY KEY (id_equipo) 
)

create table pacientes(
	id_paciente int IDENTITY(1,1),
	nombre varchar(25),
	apellido varchar(25),
	cedula varchar(25),
	edad int,
	genero varchar(10),
	ubicacion varchar(50),
	celular varchar(25),
	correo varchar(25),
	estado varchar(10),
	id_equipo int,
	PRIMARY KEY (id_paciente),
	FOREIGN KEY (id_equipo) references equipos(id_equipo)
)

create table integrantes(
	id_integrante int IDENTITY(1,1),
	nombre varchar(25),
	apellido varchar(25),
	cedula varchar(25),
	id_equipo int,
	PRIMARY KEY (id_integrante),
	FOREIGN KEY (id_equipo) REFERENCES equipos(id_equipo)
)

insert into equipos values
	('San Miguel'),('Chepo'),('Chimán'),('Betania'),
	('El Chorrillo'),('Tocumen'),('Belisario Porras'),('Omar Torrijos'),
	('Taboga')

insert into integrantes (nombre,apellido,cedula,id_equipo) values
	('Ricardo','Ye','8-1018-2065',5),
	('Javier','Arrue','8-941-1079',4),
	('Franklin','Alvarado','8-936-2210',1),
	('Rocio','Ñañez','E-8-114992',2),
	('Carlos','Gonzalez','8-142-1432',3),
	('Camila','Torres','8-123-1567',6),
	('Joe','Rogan','8-321-1895',7),
	('Jeremy','Corbell','8-321-1755',8),
	('Bob','Lazar','8-132-6543',9)

select*from equipos
select*from integrantes
select*from pacientes

Go
create procedure insertarPaciente @nombre varchar(25), @apellido varchar(25), @cedula varchar(25), @edad int, @genero varchar(25), @ubicacion varchar(50), @celular varchar(25), @correo varchar(25), @estado varchar(20), @id_equipo int
	as 
	INSERT INTO [dbo].[pacientes] ([nombre], [apellido], [cedula], [edad], [genero], [ubicacion], [celular], [correo], [estado], [id_equipo]) VALUES (@nombre, @apellido, @cedula, @edad, @genero, @ubicacion, @celular, @correo, @estado, @id_equipo);






SELECT id_paciente, nombre, apellido, cedula, edad, genero, ubicacion, celular, correo, estado, id_equipo FROM pacientes WHERE (id_paciente = SCOPE_IDENTITY())