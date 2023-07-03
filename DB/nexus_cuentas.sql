-- phpMyAdmin SQL Dump
-- version 5.0.4deb2+deb11u1
-- https://www.phpmyadmin.net/
--
-- Servidor: localhost:3306
-- Tiempo de generación: 02-07-2023 a las 22:07:16
-- Versión del servidor: 10.5.19-MariaDB-0+deb11u2
-- Versión de PHP: 7.4.33

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
START TRANSACTION;
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;

--
-- Base de datos: `nexus_cuentas`
--
CREATE DATABASE IF NOT EXISTS `nexus_cuentas` DEFAULT CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci;
USE `nexus_cuentas`;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `cuentas`
--

CREATE TABLE `cuentas` (
  `id` mediumint(8) UNSIGNED NOT NULL,
  `username` varchar(24) NOT NULL,
  `email` varchar(64) NOT NULL,
  `password` varchar(64) NOT NULL,
  `salt` varchar(32) NOT NULL,
  `id_recuperacion` varchar(32) DEFAULT NULL,
  `date_created` timestamp NULL DEFAULT current_timestamp(),
  `last_ip` varchar(16) DEFAULT NULL,
  `date_last_login` timestamp NULL DEFAULT current_timestamp(),
  `gemas` int(12) DEFAULT 0,
  `status` tinyint(1) NOT NULL DEFAULT 0,
  `id_confirmacion` varchar(128) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Volcado de datos para la tabla `cuentas`
--

INSERT INTO `cuentas` (`id`, `username`, `email`, `password`, `salt`, `id_recuperacion`, `date_created`, `last_ip`, `date_last_login`, `gemas`, `status`, `id_confirmacion`) VALUES
(1, 'Lorwik', 'asd@asd.com', '93960fd04cc56c4803a6ee4e50f41e04ace49cd5c7b7b5f5eb255a6921f8a8f4', 'Qgyq9NDa#0z!p1GAY~vbJRCRClmkGtns', NULL, '2023-04-23 12:15:38', '127.0.0.1', '2023-07-02 06:16:52', 0, 1, 'VERIFICADA'),
(2, 'test', 'test@test.com', '5f288f29813690563e614d74c891246bd6db0a8eefb2a88c702e8be7e07ed263', 'rb$x)k_i@QkDh-ng(gmvd6ju0b_ON95u', NULL, '2023-05-26 06:10:30', '127.0.0.1', '2023-05-28 19:48:56', 0, 1, 'VERIFICADA');

--
-- Índices para tablas volcadas
--

--
-- Indices de la tabla `cuentas`
--
ALTER TABLE `cuentas`
  ADD PRIMARY KEY (`id`);

--
-- AUTO_INCREMENT de las tablas volcadas
--

--
-- AUTO_INCREMENT de la tabla `cuentas`
--
ALTER TABLE `cuentas`
  MODIFY `id` mediumint(8) UNSIGNED NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=3;
COMMIT;

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
