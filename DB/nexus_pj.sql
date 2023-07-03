-- phpMyAdmin SQL Dump
-- version 5.0.4deb2+deb11u1
-- https://www.phpmyadmin.net/
--
-- Servidor: localhost:3306
-- Tiempo de generación: 02-07-2023 a las 22:07:33
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
-- Base de datos: `nexus_pj`
--
CREATE DATABASE IF NOT EXISTS `nexus_pj` DEFAULT CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci;
USE `nexus_pj`;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `atributos`
--

CREATE TABLE `atributos` (
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `att1` tinyint(3) UNSIGNED NOT NULL,
  `att2` tinyint(3) UNSIGNED NOT NULL,
  `att3` tinyint(3) UNSIGNED NOT NULL,
  `att4` tinyint(3) UNSIGNED NOT NULL,
  `att5` tinyint(3) UNSIGNED NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Volcado de datos para la tabla `atributos`
--

INSERT INTO `atributos` (`user_id`, `att1`, `att2`, `att3`, `att4`, `att5`) VALUES
(1, 19, 19, 18, 18, 20),
(2, 19, 19, 18, 18, 20);

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `banco_items`
--

CREATE TABLE `banco_items` (
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `item_id1` smallint(5) UNSIGNED DEFAULT 0,
  `amount1` smallint(5) UNSIGNED DEFAULT 0,
  `item_id2` smallint(5) UNSIGNED DEFAULT 0,
  `amount2` smallint(5) UNSIGNED DEFAULT 0,
  `item_id3` smallint(5) UNSIGNED DEFAULT 0,
  `amount3` smallint(5) UNSIGNED DEFAULT 0,
  `item_id4` smallint(5) UNSIGNED DEFAULT 0,
  `amount4` smallint(5) UNSIGNED DEFAULT 0,
  `item_id5` smallint(5) UNSIGNED DEFAULT 0,
  `amount5` smallint(5) UNSIGNED DEFAULT 0,
  `item_id6` smallint(5) UNSIGNED DEFAULT 0,
  `amount6` smallint(5) UNSIGNED DEFAULT 0,
  `item_id7` smallint(5) UNSIGNED DEFAULT 0,
  `amount7` smallint(5) UNSIGNED DEFAULT 0,
  `item_id8` smallint(5) UNSIGNED DEFAULT 0,
  `amount8` smallint(5) UNSIGNED DEFAULT 0,
  `item_id9` smallint(5) UNSIGNED DEFAULT 0,
  `amount9` smallint(5) UNSIGNED DEFAULT 0,
  `item_id10` smallint(5) UNSIGNED DEFAULT 0,
  `amount10` smallint(5) UNSIGNED DEFAULT 0,
  `item_id11` smallint(5) UNSIGNED DEFAULT 0,
  `amount11` smallint(5) UNSIGNED DEFAULT 0,
  `item_id12` smallint(5) UNSIGNED DEFAULT 0,
  `amount12` smallint(5) UNSIGNED DEFAULT 0,
  `item_id13` smallint(5) UNSIGNED DEFAULT 0,
  `amount13` smallint(5) UNSIGNED DEFAULT 0,
  `item_id14` smallint(5) UNSIGNED DEFAULT 0,
  `amount14` smallint(5) UNSIGNED DEFAULT 0,
  `item_id15` smallint(5) UNSIGNED DEFAULT 0,
  `amount15` smallint(5) UNSIGNED DEFAULT 0,
  `item_id16` smallint(5) UNSIGNED DEFAULT 0,
  `amount16` smallint(5) UNSIGNED DEFAULT 0,
  `item_id17` smallint(5) UNSIGNED DEFAULT 0,
  `amount17` smallint(5) UNSIGNED DEFAULT 0,
  `item_id18` smallint(5) UNSIGNED DEFAULT 0,
  `amount18` smallint(5) UNSIGNED DEFAULT 0,
  `item_id19` smallint(5) UNSIGNED DEFAULT 0,
  `amount19` smallint(5) UNSIGNED DEFAULT 0,
  `item_id20` smallint(5) UNSIGNED DEFAULT 0,
  `amount20` smallint(5) UNSIGNED DEFAULT 0,
  `item_id21` smallint(5) UNSIGNED DEFAULT 0,
  `amount21` smallint(5) UNSIGNED DEFAULT 0,
  `item_id22` smallint(5) UNSIGNED DEFAULT 0,
  `amount22` smallint(5) UNSIGNED DEFAULT 0,
  `item_id23` smallint(5) UNSIGNED DEFAULT 0,
  `amount23` smallint(5) UNSIGNED DEFAULT 0,
  `item_id24` smallint(5) UNSIGNED DEFAULT 0,
  `amount24` smallint(5) UNSIGNED DEFAULT 0,
  `item_id25` smallint(5) UNSIGNED DEFAULT 0,
  `amount25` smallint(5) UNSIGNED DEFAULT 0,
  `item_id26` smallint(5) UNSIGNED DEFAULT 0,
  `amount26` smallint(5) UNSIGNED DEFAULT 0,
  `item_id27` smallint(5) UNSIGNED DEFAULT 0,
  `amount27` smallint(5) UNSIGNED DEFAULT 0,
  `item_id28` smallint(5) UNSIGNED DEFAULT 0,
  `amount28` smallint(5) UNSIGNED DEFAULT 0,
  `item_id29` smallint(5) UNSIGNED DEFAULT 0,
  `amount29` smallint(5) UNSIGNED DEFAULT 0,
  `item_id30` smallint(5) UNSIGNED DEFAULT 0,
  `amount30` smallint(5) UNSIGNED DEFAULT 0,
  `item_id31` smallint(5) UNSIGNED DEFAULT 0,
  `amount31` smallint(5) UNSIGNED DEFAULT 0,
  `item_id32` smallint(5) UNSIGNED DEFAULT 0,
  `amount32` smallint(5) UNSIGNED DEFAULT 0,
  `item_id33` smallint(5) UNSIGNED DEFAULT 0,
  `amount33` smallint(5) UNSIGNED DEFAULT 0,
  `item_id34` smallint(5) UNSIGNED DEFAULT 0,
  `amount34` smallint(5) UNSIGNED DEFAULT 0,
  `item_id35` smallint(5) UNSIGNED DEFAULT 0,
  `amount35` smallint(5) UNSIGNED DEFAULT 0,
  `item_id36` smallint(5) UNSIGNED DEFAULT 0,
  `amount36` smallint(5) UNSIGNED DEFAULT 0,
  `item_id37` smallint(5) UNSIGNED DEFAULT 0,
  `amount37` smallint(5) UNSIGNED DEFAULT 0,
  `item_id38` smallint(5) UNSIGNED DEFAULT 0,
  `amount38` smallint(5) UNSIGNED DEFAULT 0,
  `item_id39` smallint(5) UNSIGNED DEFAULT 0,
  `amount39` smallint(5) UNSIGNED DEFAULT 0,
  `item_id40` smallint(5) UNSIGNED DEFAULT 0,
  `amount40` smallint(5) UNSIGNED DEFAULT 0
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Volcado de datos para la tabla `banco_items`
--

INSERT INTO `banco_items` (`user_id`, `item_id1`, `amount1`, `item_id2`, `amount2`, `item_id3`, `amount3`, `item_id4`, `amount4`, `item_id5`, `amount5`, `item_id6`, `amount6`, `item_id7`, `amount7`, `item_id8`, `amount8`, `item_id9`, `amount9`, `item_id10`, `amount10`, `item_id11`, `amount11`, `item_id12`, `amount12`, `item_id13`, `amount13`, `item_id14`, `amount14`, `item_id15`, `amount15`, `item_id16`, `amount16`, `item_id17`, `amount17`, `item_id18`, `amount18`, `item_id19`, `amount19`, `item_id20`, `amount20`, `item_id21`, `amount21`, `item_id22`, `amount22`, `item_id23`, `amount23`, `item_id24`, `amount24`, `item_id25`, `amount25`, `item_id26`, `amount26`, `item_id27`, `amount27`, `item_id28`, `amount28`, `item_id29`, `amount29`, `item_id30`, `amount30`, `item_id31`, `amount31`, `item_id32`, `amount32`, `item_id33`, `amount33`, `item_id34`, `amount34`, `item_id35`, `amount35`, `item_id36`, `amount36`, `item_id37`, `amount37`, `item_id38`, `amount38`, `item_id39`, `amount39`, `item_id40`, `amount40`) VALUES
(1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0),
(2, 138, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `familiar`
--

CREATE TABLE `familiar` (
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `nombre` varchar(30) NOT NULL,
  `level` smallint(5) UNSIGNED NOT NULL,
  `exp` int(10) UNSIGNED NOT NULL,
  `elu` int(10) UNSIGNED NOT NULL,
  `tipo` int(4) UNSIGNED NOT NULL,
  `min_hp` smallint(5) UNSIGNED NOT NULL,
  `max_hp` smallint(5) UNSIGNED NOT NULL,
  `min_hit` smallint(5) UNSIGNED NOT NULL,
  `max_hit` smallint(5) UNSIGNED NOT NULL,
  `h_id1` smallint(5) UNSIGNED DEFAULT 0,
  `h_id2` smallint(5) UNSIGNED DEFAULT 0,
  `h_id3` smallint(5) UNSIGNED DEFAULT 0,
  `h_id4` smallint(5) UNSIGNED DEFAULT 0
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Volcado de datos para la tabla `familiar`
--

INSERT INTO `familiar` (`user_id`, `nombre`, `level`, `exp`, `elu`, `tipo`, `min_hp`, `max_hp`, `min_hit`, `max_hit`, `h_id1`, `h_id2`, `h_id3`, `h_id4`) VALUES
(1, '', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0),
(2, '', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `inventario_items`
--

CREATE TABLE `inventario_items` (
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `item_id1` smallint(5) UNSIGNED DEFAULT 0,
  `amount1` smallint(5) UNSIGNED DEFAULT 0,
  `is_equipped1` tinyint(1) UNSIGNED DEFAULT 0,
  `item_id2` smallint(5) UNSIGNED DEFAULT 0,
  `amount2` smallint(5) UNSIGNED DEFAULT 0,
  `is_equipped2` tinyint(1) UNSIGNED DEFAULT 0,
  `item_id3` smallint(5) UNSIGNED DEFAULT 0,
  `amount3` smallint(5) UNSIGNED DEFAULT 0,
  `is_equipped3` tinyint(1) UNSIGNED DEFAULT 0,
  `item_id4` smallint(5) UNSIGNED DEFAULT 0,
  `amount4` smallint(5) UNSIGNED DEFAULT 0,
  `is_equipped4` tinyint(1) UNSIGNED DEFAULT 0,
  `item_id5` smallint(5) UNSIGNED DEFAULT 0,
  `amount5` smallint(5) UNSIGNED DEFAULT 0,
  `is_equipped5` tinyint(1) UNSIGNED DEFAULT 0,
  `item_id6` smallint(5) UNSIGNED DEFAULT 0,
  `amount6` smallint(5) UNSIGNED DEFAULT 0,
  `is_equipped6` tinyint(1) UNSIGNED DEFAULT 0,
  `item_id7` smallint(5) UNSIGNED DEFAULT 0,
  `amount7` smallint(5) UNSIGNED DEFAULT 0,
  `is_equipped7` tinyint(1) UNSIGNED DEFAULT 0,
  `item_id8` smallint(5) UNSIGNED DEFAULT 0,
  `amount8` smallint(5) UNSIGNED DEFAULT 0,
  `is_equipped8` tinyint(1) UNSIGNED DEFAULT 0,
  `item_id9` smallint(5) UNSIGNED DEFAULT 0,
  `amount9` smallint(5) UNSIGNED DEFAULT 0,
  `is_equipped9` tinyint(1) UNSIGNED DEFAULT 0,
  `item_id10` smallint(5) UNSIGNED DEFAULT 0,
  `amount10` smallint(5) UNSIGNED DEFAULT 0,
  `is_equipped10` tinyint(1) UNSIGNED DEFAULT 0,
  `item_id11` smallint(5) UNSIGNED DEFAULT 0,
  `amount11` smallint(5) UNSIGNED DEFAULT 0,
  `is_equipped11` tinyint(1) UNSIGNED DEFAULT 0,
  `item_id12` smallint(5) UNSIGNED DEFAULT 0,
  `amount12` smallint(5) UNSIGNED DEFAULT 0,
  `is_equipped12` tinyint(1) UNSIGNED DEFAULT 0,
  `item_id13` smallint(5) UNSIGNED DEFAULT 0,
  `amount13` smallint(5) UNSIGNED DEFAULT 0,
  `is_equipped13` tinyint(1) UNSIGNED DEFAULT 0,
  `item_id14` smallint(5) UNSIGNED DEFAULT 0,
  `amount14` smallint(5) UNSIGNED DEFAULT 0,
  `is_equipped14` tinyint(1) UNSIGNED DEFAULT 0,
  `item_id15` smallint(5) UNSIGNED DEFAULT 0,
  `amount15` smallint(5) UNSIGNED DEFAULT 0,
  `is_equipped15` tinyint(1) UNSIGNED DEFAULT 0,
  `item_id16` smallint(5) UNSIGNED DEFAULT 0,
  `amount16` smallint(5) UNSIGNED DEFAULT 0,
  `is_equipped16` tinyint(1) UNSIGNED DEFAULT 0,
  `item_id17` smallint(5) UNSIGNED DEFAULT 0,
  `amount17` smallint(5) UNSIGNED DEFAULT 0,
  `is_equipped17` tinyint(1) UNSIGNED DEFAULT 0,
  `item_id18` smallint(5) UNSIGNED DEFAULT 0,
  `amount18` smallint(5) UNSIGNED DEFAULT 0,
  `is_equipped18` tinyint(1) UNSIGNED DEFAULT 0,
  `item_id19` smallint(5) UNSIGNED DEFAULT 0,
  `amount19` smallint(5) UNSIGNED DEFAULT 0,
  `is_equipped19` tinyint(1) UNSIGNED DEFAULT 0,
  `item_id20` smallint(5) UNSIGNED DEFAULT 0,
  `amount20` smallint(5) UNSIGNED DEFAULT 0,
  `is_equipped20` tinyint(1) UNSIGNED DEFAULT 0,
  `item_id21` smallint(5) UNSIGNED DEFAULT 0,
  `amount21` smallint(5) UNSIGNED DEFAULT 0,
  `is_equipped21` tinyint(1) UNSIGNED DEFAULT 0,
  `item_id22` smallint(5) UNSIGNED DEFAULT 0,
  `amount22` smallint(5) UNSIGNED DEFAULT 0,
  `is_equipped22` tinyint(1) UNSIGNED DEFAULT 0,
  `item_id23` smallint(5) UNSIGNED DEFAULT 0,
  `amount23` smallint(5) UNSIGNED DEFAULT 0,
  `is_equipped23` tinyint(1) UNSIGNED DEFAULT 0,
  `item_id24` smallint(5) UNSIGNED DEFAULT 0,
  `amount24` smallint(5) UNSIGNED DEFAULT 0,
  `is_equipped24` tinyint(1) UNSIGNED DEFAULT 0,
  `item_id25` smallint(5) UNSIGNED DEFAULT 0,
  `amount25` smallint(5) UNSIGNED DEFAULT 0,
  `is_equipped25` tinyint(1) UNSIGNED DEFAULT 0,
  `item_id26` smallint(5) UNSIGNED DEFAULT 0,
  `amount26` smallint(5) UNSIGNED DEFAULT 0,
  `is_equipped26` tinyint(1) UNSIGNED DEFAULT 0,
  `item_id27` smallint(5) UNSIGNED DEFAULT 0,
  `amount27` smallint(5) UNSIGNED DEFAULT 0,
  `is_equipped27` tinyint(1) UNSIGNED DEFAULT 0,
  `item_id28` smallint(5) UNSIGNED DEFAULT 0,
  `amount28` smallint(5) UNSIGNED DEFAULT 0,
  `is_equipped28` tinyint(1) UNSIGNED DEFAULT 0,
  `item_id29` smallint(5) UNSIGNED DEFAULT 0,
  `amount29` smallint(5) UNSIGNED DEFAULT 0,
  `is_equipped29` tinyint(1) UNSIGNED DEFAULT 0,
  `item_id30` smallint(5) UNSIGNED DEFAULT 0,
  `amount30` smallint(5) UNSIGNED DEFAULT 0,
  `is_equipped30` tinyint(1) UNSIGNED DEFAULT 0,
  `item_id31` smallint(5) UNSIGNED DEFAULT 0,
  `amount31` smallint(5) UNSIGNED DEFAULT 0,
  `is_equipped31` tinyint(1) UNSIGNED DEFAULT 0,
  `item_id32` smallint(5) UNSIGNED DEFAULT 0,
  `amount32` smallint(5) UNSIGNED DEFAULT 0,
  `is_equipped32` tinyint(1) UNSIGNED DEFAULT 0,
  `item_id33` smallint(5) UNSIGNED DEFAULT 0,
  `amount33` smallint(5) UNSIGNED DEFAULT 0,
  `is_equipped33` tinyint(1) UNSIGNED DEFAULT 0,
  `item_id34` smallint(5) UNSIGNED DEFAULT 0,
  `amount34` smallint(5) UNSIGNED DEFAULT 0,
  `is_equipped34` tinyint(1) UNSIGNED DEFAULT 0,
  `item_id35` smallint(5) UNSIGNED DEFAULT 0,
  `amount35` smallint(5) UNSIGNED DEFAULT 0,
  `is_equipped35` tinyint(1) UNSIGNED DEFAULT 0
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Volcado de datos para la tabla `inventario_items`
--

INSERT INTO `inventario_items` (`user_id`, `item_id1`, `amount1`, `is_equipped1`, `item_id2`, `amount2`, `is_equipped2`, `item_id3`, `amount3`, `is_equipped3`, `item_id4`, `amount4`, `is_equipped4`, `item_id5`, `amount5`, `is_equipped5`, `item_id6`, `amount6`, `is_equipped6`, `item_id7`, `amount7`, `is_equipped7`, `item_id8`, `amount8`, `is_equipped8`, `item_id9`, `amount9`, `is_equipped9`, `item_id10`, `amount10`, `is_equipped10`, `item_id11`, `amount11`, `is_equipped11`, `item_id12`, `amount12`, `is_equipped12`, `item_id13`, `amount13`, `is_equipped13`, `item_id14`, `amount14`, `is_equipped14`, `item_id15`, `amount15`, `is_equipped15`, `item_id16`, `amount16`, `is_equipped16`, `item_id17`, `amount17`, `is_equipped17`, `item_id18`, `amount18`, `is_equipped18`, `item_id19`, `amount19`, `is_equipped19`, `item_id20`, `amount20`, `is_equipped20`, `item_id21`, `amount21`, `is_equipped21`, `item_id22`, `amount22`, `is_equipped22`, `item_id23`, `amount23`, `is_equipped23`, `item_id24`, `amount24`, `is_equipped24`, `item_id25`, `amount25`, `is_equipped25`, `item_id26`, `amount26`, `is_equipped26`, `item_id27`, `amount27`, `is_equipped27`, `item_id28`, `amount28`, `is_equipped28`, `item_id29`, `amount29`, `is_equipped29`, `item_id30`, `amount30`, `is_equipped30`, `item_id31`, `amount31`, `is_equipped31`, `item_id32`, `amount32`, `is_equipped32`, `item_id33`, `amount33`, `is_equipped33`, `item_id34`, `amount34`, `is_equipped34`, `item_id35`, `amount35`, `is_equipped35`) VALUES
(1, 391, 1, 1, 759, 1, 1, 187, 1, 0, 198, 1, 0, 192, 10, 0, 58, 15, 0, 0, 0, 0, 58, 9992, 0, 0, 0, 0, 0, 0, 0, 163, 11, 0, 127, 1, 1, 138, 1, 0, 139, 3, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 161, 93, 0, 64, 10, 0, 824, 1, 0, 474, 1, 0, 779, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0),
(2, 515, 1, 1, 759, 1, 0, 404, 1, 1, 198, 1, 0, 127, 1, 1, 709, 28, 0, 708, 44, 0, 58, 4099, 0, 163, 62, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 26, 83, 0, 161, 49, 0, 779, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `personaje`
--

CREATE TABLE `personaje` (
  `id` mediumint(8) UNSIGNED NOT NULL,
  `cuenta_id` mediumint(8) UNSIGNED NOT NULL,
  `deleted` tinyint(1) NOT NULL DEFAULT 0,
  `name` varchar(30) NOT NULL,
  `level` smallint(5) UNSIGNED NOT NULL,
  `exp` int(10) UNSIGNED NOT NULL,
  `elu` int(10) UNSIGNED NOT NULL,
  `elo` int(10) NOT NULL,
  `genre_id` tinyint(3) UNSIGNED NOT NULL,
  `race_id` tinyint(3) UNSIGNED NOT NULL,
  `class_id` tinyint(3) UNSIGNED NOT NULL,
  `home_id` tinyint(3) UNSIGNED NOT NULL,
  `description` varchar(255) DEFAULT NULL,
  `gold` int(10) UNSIGNED NOT NULL,
  `bank_gold` int(10) UNSIGNED NOT NULL DEFAULT 0,
  `pet_amount` tinyint(3) UNSIGNED NOT NULL DEFAULT 0,
  `votes_amount` smallint(5) UNSIGNED DEFAULT 0,
  `pos_map` smallint(5) UNSIGNED NOT NULL,
  `pos_x` tinyint(3) UNSIGNED NOT NULL,
  `pos_y` tinyint(3) UNSIGNED NOT NULL,
  `last_map` tinyint(3) UNSIGNED NOT NULL DEFAULT 1,
  `body_id` smallint(5) UNSIGNED NOT NULL,
  `head_id` smallint(5) UNSIGNED NOT NULL,
  `weapon_id` smallint(5) UNSIGNED NOT NULL,
  `helmet_id` smallint(5) UNSIGNED NOT NULL,
  `shield_id` smallint(5) UNSIGNED NOT NULL,
  `aura_id` int(24) DEFAULT 0,
  `aura_color` int(24) DEFAULT 0,
  `heading` tinyint(3) UNSIGNED NOT NULL DEFAULT 3,
  `items_amount` tinyint(3) UNSIGNED NOT NULL,
  `slot_armour` tinyint(3) UNSIGNED DEFAULT NULL,
  `slot_weapon` tinyint(3) UNSIGNED DEFAULT NULL,
  `slot_nudillos` tinyint(3) UNSIGNED DEFAULT NULL,
  `slot_helmet` tinyint(3) UNSIGNED DEFAULT NULL,
  `slot_shield` tinyint(3) UNSIGNED DEFAULT NULL,
  `slot_ammo` tinyint(3) UNSIGNED DEFAULT NULL,
  `slot_ship` tinyint(3) UNSIGNED DEFAULT NULL,
  `slot_ring` tinyint(3) UNSIGNED DEFAULT NULL,
  `slot_bag` tinyint(3) UNSIGNED DEFAULT NULL,
  `min_hp` smallint(5) UNSIGNED NOT NULL,
  `max_hp` smallint(5) UNSIGNED NOT NULL,
  `min_man` smallint(5) UNSIGNED NOT NULL,
  `max_man` smallint(5) UNSIGNED NOT NULL,
  `min_sta` smallint(5) UNSIGNED NOT NULL,
  `max_sta` smallint(5) UNSIGNED NOT NULL,
  `min_ham` smallint(5) UNSIGNED NOT NULL,
  `max_ham` smallint(5) UNSIGNED NOT NULL,
  `min_sed` smallint(5) UNSIGNED NOT NULL,
  `max_sed` smallint(5) UNSIGNED NOT NULL,
  `min_hit` smallint(5) UNSIGNED NOT NULL,
  `max_hit` smallint(5) UNSIGNED NOT NULL,
  `killed_npcs` smallint(5) UNSIGNED NOT NULL DEFAULT 0,
  `killed` smallint(5) UNSIGNED NOT NULL DEFAULT 0,
  `rep_asesino` mediumint(8) UNSIGNED NOT NULL DEFAULT 0,
  `rep_bandido` mediumint(8) UNSIGNED NOT NULL DEFAULT 0,
  `rep_burgues` mediumint(8) UNSIGNED NOT NULL DEFAULT 0,
  `rep_ladron` mediumint(8) UNSIGNED NOT NULL DEFAULT 0,
  `rep_noble` mediumint(8) UNSIGNED NOT NULL,
  `rep_plebe` mediumint(8) UNSIGNED NOT NULL,
  `rep_average` mediumint(9) NOT NULL,
  `is_naked` tinyint(1) NOT NULL DEFAULT 0,
  `is_poisoned` tinyint(1) NOT NULL DEFAULT 0,
  `is_incinerado` tinyint(1) DEFAULT 0,
  `is_hidden` tinyint(1) NOT NULL DEFAULT 0,
  `is_hungry` tinyint(1) NOT NULL DEFAULT 0,
  `is_thirsty` tinyint(1) NOT NULL DEFAULT 0,
  `is_ban` tinyint(1) NOT NULL DEFAULT 0,
  `is_dead` tinyint(1) NOT NULL DEFAULT 0,
  `is_sailing` tinyint(1) NOT NULL DEFAULT 0,
  `is_paralyzed` tinyint(1) NOT NULL DEFAULT 0,
  `is_logged` tinyint(1) NOT NULL DEFAULT 0,
  `counter_pena` smallint(5) UNSIGNED NOT NULL DEFAULT 0,
  `counter_connected` int(10) UNSIGNED NOT NULL DEFAULT 0,
  `counter_training` int(10) UNSIGNED NOT NULL DEFAULT 0,
  `pertenece_consejo_real` tinyint(1) NOT NULL DEFAULT 0,
  `pertenece_consejo_caos` tinyint(1) NOT NULL DEFAULT 0,
  `pertenece_real` tinyint(1) NOT NULL DEFAULT 0,
  `pertenece_caos` tinyint(1) NOT NULL DEFAULT 0,
  `ciudadanos_matados` smallint(5) UNSIGNED NOT NULL DEFAULT 0,
  `criminales_matados` smallint(5) UNSIGNED NOT NULL DEFAULT 0,
  `recibio_armadura_real` tinyint(1) NOT NULL DEFAULT 0,
  `recibio_armadura_caos` tinyint(1) NOT NULL DEFAULT 0,
  `recibio_exp_real` tinyint(1) NOT NULL DEFAULT 0,
  `recibio_exp_caos` tinyint(1) NOT NULL DEFAULT 0,
  `recompensas_real` tinyint(3) UNSIGNED DEFAULT 0,
  `recompensas_caos` tinyint(3) UNSIGNED DEFAULT 0,
  `reenlistadas` smallint(5) UNSIGNED NOT NULL DEFAULT 0,
  `fecha_ingreso` timestamp NOT NULL DEFAULT current_timestamp() ON UPDATE current_timestamp(),
  `nivel_ingreso` smallint(5) UNSIGNED DEFAULT NULL,
  `matados_ingreso` smallint(5) UNSIGNED DEFAULT NULL,
  `siguiente_recompensa` smallint(5) UNSIGNED DEFAULT NULL,
  `guild_index` smallint(5) UNSIGNED DEFAULT 0,
  `guild_aspirant_index` smallint(5) UNSIGNED DEFAULT NULL,
  `guild_member_history` varchar(1024) DEFAULT NULL,
  `guild_requests_history` varchar(1024) DEFAULT NULL,
  `guild_rejected_because` varchar(255) DEFAULT NULL,
  `is_global` tinyint(1) DEFAULT 1,
  `modocombate` tinyint(4) DEFAULT 0,
  `seguro` tinyint(1) DEFAULT 0,
  `pareja` varchar(30) DEFAULT '',
  `profesionA` int(2) NOT NULL DEFAULT 0,
  `profesionB` int(2) NOT NULL DEFAULT 0
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Volcado de datos para la tabla `personaje`
--

INSERT INTO `personaje` (`id`, `cuenta_id`, `deleted`, `name`, `level`, `exp`, `elu`, `elo`, `genre_id`, `race_id`, `class_id`, `home_id`, `description`, `gold`, `bank_gold`, `pet_amount`, `votes_amount`, `pos_map`, `pos_x`, `pos_y`, `last_map`, `body_id`, `head_id`, `weapon_id`, `helmet_id`, `shield_id`, `aura_id`, `aura_color`, `heading`, `items_amount`, `slot_armour`, `slot_weapon`, `slot_nudillos`, `slot_helmet`, `slot_shield`, `slot_ammo`, `slot_ship`, `slot_ring`, `slot_bag`, `min_hp`, `max_hp`, `min_man`, `max_man`, `min_sta`, `max_sta`, `min_ham`, `max_ham`, `min_sed`, `max_sed`, `min_hit`, `max_hit`, `killed_npcs`, `killed`, `rep_asesino`, `rep_bandido`, `rep_burgues`, `rep_ladron`, `rep_noble`, `rep_plebe`, `rep_average`, `is_naked`, `is_poisoned`, `is_incinerado`, `is_hidden`, `is_hungry`, `is_thirsty`, `is_ban`, `is_dead`, `is_sailing`, `is_paralyzed`, `is_logged`, `counter_pena`, `counter_connected`, `counter_training`, `pertenece_consejo_real`, `pertenece_consejo_caos`, `pertenece_real`, `pertenece_caos`, `ciudadanos_matados`, `criminales_matados`, `recibio_armadura_real`, `recibio_armadura_caos`, `recibio_exp_real`, `recibio_exp_caos`, `recompensas_real`, `recompensas_caos`, `reenlistadas`, `fecha_ingreso`, `nivel_ingreso`, `matados_ingreso`, `siguiente_recompensa`, `guild_index`, `guild_aspirant_index`, `guild_member_history`, `guild_requests_history`, `guild_rejected_because`, `is_global`, `modocombate`, `seguro`, `pareja`, `profesionA`, `profesionB`) VALUES
(1, 1, 0, 'Lorwik', 32, 88426, 757662, 1000, 1, 1, 9, 2, '', 369868, 0, 0, 0, 34, 29, 73, 1, 45, 4, 19, 2, 2, 0, 0, 2, 19, 1, 2, 0, 0, 0, 0, 0, 0, NULL, 286, 286, 608, 608, 505, 505, 100, 100, 100, 100, 94, 95, 17, 0, 0, 0, 0, 0, 9500, 134, 1606, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1595, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '2023-07-01 20:03:15', 0, 0, 0, 0, NULL, NULL, NULL, '', 1, 1, 0, '', 19, 20),
(2, 1, 0, 'Erwin', 22, 136877, 180334, 1000, 1, 1, 9, 2, '', 5779, 0, 0, 0, 34, 40, 12, 1, 126, 4, 0, 2, 6, 0, 0, 3, 14, 1, 5, 0, 0, 3, 0, 0, 0, NULL, 195, 195, 428, 428, 375, 375, 60, 100, 60, 100, 64, 65, 288, 2, 15000, 22000, 0, 0, 137500, 9546, 18341, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 7373, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '2023-07-02 06:16:52', 0, 0, 0, 0, NULL, NULL, NULL, '', 1, 1, 0, '', 23, 20);

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `pet`
--

CREATE TABLE `pet` (
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `pet1` smallint(5) UNSIGNED DEFAULT NULL,
  `pet2` smallint(5) UNSIGNED DEFAULT NULL,
  `pet3` smallint(5) UNSIGNED DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Volcado de datos para la tabla `pet`
--

INSERT INTO `pet` (`user_id`, `pet1`, `pet2`, `pet3`) VALUES
(1, 0, 0, 0),
(2, 0, 0, 0);

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `profesion_primaria`
--

CREATE TABLE `profesion_primaria` (
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `profesion` smallint(4) UNSIGNED NOT NULL DEFAULT 0,
  `receta1` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta2` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta3` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta4` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta5` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta6` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta7` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta8` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta9` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta10` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta11` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta12` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta13` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta14` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta15` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta16` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta17` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta18` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta19` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta20` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta21` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta22` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta23` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta24` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta25` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta26` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta27` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta28` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta29` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta30` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta31` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta32` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta33` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta34` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta35` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta36` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta37` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta38` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta39` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta40` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta41` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta42` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta43` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta44` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta45` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta46` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta47` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta48` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta49` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta50` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta51` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta52` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta53` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta54` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta55` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta56` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta57` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta58` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta59` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta60` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta61` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta62` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta63` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta64` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta65` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta66` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta67` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta68` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta69` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta70` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta71` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta72` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta73` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta74` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta75` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta76` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta77` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta78` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta79` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta80` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta81` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta82` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta83` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta84` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta85` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta86` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta87` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta88` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta89` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta90` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta91` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta92` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta93` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta94` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta95` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta96` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta97` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta98` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta99` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta100` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta101` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta102` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta103` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta104` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta105` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta106` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta107` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta108` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta109` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta110` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta111` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta112` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta113` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta114` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta115` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta116` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta117` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta118` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta119` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta120` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta121` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta122` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta123` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta124` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta125` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta126` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta127` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta128` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta129` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta130` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta131` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta132` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta133` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta134` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta135` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta136` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta137` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta138` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta139` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta140` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta141` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta142` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta143` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta144` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta145` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta146` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta147` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta148` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta149` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta150` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta151` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta152` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta153` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta154` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta155` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta156` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta157` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta158` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta159` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta160` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta161` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta162` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta163` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta164` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta165` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta166` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta167` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta168` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta169` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta170` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta171` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta172` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta173` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta174` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta175` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta176` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta177` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta178` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta179` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta180` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta181` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta182` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta183` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta184` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta185` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta186` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta187` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta188` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta189` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta190` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta191` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta192` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta193` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta194` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta195` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta196` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta197` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta198` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta199` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta200` smallint(6) UNSIGNED NOT NULL DEFAULT 0
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Volcado de datos para la tabla `profesion_primaria`
--

INSERT INTO `profesion_primaria` (`user_id`, `profesion`, `receta1`, `receta2`, `receta3`, `receta4`, `receta5`, `receta6`, `receta7`, `receta8`, `receta9`, `receta10`, `receta11`, `receta12`, `receta13`, `receta14`, `receta15`, `receta16`, `receta17`, `receta18`, `receta19`, `receta20`, `receta21`, `receta22`, `receta23`, `receta24`, `receta25`, `receta26`, `receta27`, `receta28`, `receta29`, `receta30`, `receta31`, `receta32`, `receta33`, `receta34`, `receta35`, `receta36`, `receta37`, `receta38`, `receta39`, `receta40`, `receta41`, `receta42`, `receta43`, `receta44`, `receta45`, `receta46`, `receta47`, `receta48`, `receta49`, `receta50`, `receta51`, `receta52`, `receta53`, `receta54`, `receta55`, `receta56`, `receta57`, `receta58`, `receta59`, `receta60`, `receta61`, `receta62`, `receta63`, `receta64`, `receta65`, `receta66`, `receta67`, `receta68`, `receta69`, `receta70`, `receta71`, `receta72`, `receta73`, `receta74`, `receta75`, `receta76`, `receta77`, `receta78`, `receta79`, `receta80`, `receta81`, `receta82`, `receta83`, `receta84`, `receta85`, `receta86`, `receta87`, `receta88`, `receta89`, `receta90`, `receta91`, `receta92`, `receta93`, `receta94`, `receta95`, `receta96`, `receta97`, `receta98`, `receta99`, `receta100`, `receta101`, `receta102`, `receta103`, `receta104`, `receta105`, `receta106`, `receta107`, `receta108`, `receta109`, `receta110`, `receta111`, `receta112`, `receta113`, `receta114`, `receta115`, `receta116`, `receta117`, `receta118`, `receta119`, `receta120`, `receta121`, `receta122`, `receta123`, `receta124`, `receta125`, `receta126`, `receta127`, `receta128`, `receta129`, `receta130`, `receta131`, `receta132`, `receta133`, `receta134`, `receta135`, `receta136`, `receta137`, `receta138`, `receta139`, `receta140`, `receta141`, `receta142`, `receta143`, `receta144`, `receta145`, `receta146`, `receta147`, `receta148`, `receta149`, `receta150`, `receta151`, `receta152`, `receta153`, `receta154`, `receta155`, `receta156`, `receta157`, `receta158`, `receta159`, `receta160`, `receta161`, `receta162`, `receta163`, `receta164`, `receta165`, `receta166`, `receta167`, `receta168`, `receta169`, `receta170`, `receta171`, `receta172`, `receta173`, `receta174`, `receta175`, `receta176`, `receta177`, `receta178`, `receta179`, `receta180`, `receta181`, `receta182`, `receta183`, `receta184`, `receta185`, `receta186`, `receta187`, `receta188`, `receta189`, `receta190`, `receta191`, `receta192`, `receta193`, `receta194`, `receta195`, `receta196`, `receta197`, `receta198`, `receta199`, `receta200`) VALUES
(1, 19, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0),
(2, 23, 163, 474, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `profesion_secundaria`
--

CREATE TABLE `profesion_secundaria` (
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `profesion` smallint(4) UNSIGNED NOT NULL DEFAULT 0,
  `receta1` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta2` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta3` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta4` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta5` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta6` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta7` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta8` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta9` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta10` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta11` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta12` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta13` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta14` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta15` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta16` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta17` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta18` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta19` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta20` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta21` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta22` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta23` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta24` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta25` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta26` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta27` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta28` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta29` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta30` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta31` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta32` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta33` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta34` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta35` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta36` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta37` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta38` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta39` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta40` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta41` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta42` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta43` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta44` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta45` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta46` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta47` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta48` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta49` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta50` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta51` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta52` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta53` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta54` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta55` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta56` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta57` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta58` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta59` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta60` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta61` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta62` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta63` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta64` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta65` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta66` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta67` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta68` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta69` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta70` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta71` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta72` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta73` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta74` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta75` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta76` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta77` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta78` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta79` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta80` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta81` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta82` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta83` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta84` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta85` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta86` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta87` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta88` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta89` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta90` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta91` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta92` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta93` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta94` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta95` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta96` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta97` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta98` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta99` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta100` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta101` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta102` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta103` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta104` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta105` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta106` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta107` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta108` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta109` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta110` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta111` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta112` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta113` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta114` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta115` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta116` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta117` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta118` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta119` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta120` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta121` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta122` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta123` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta124` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta125` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta126` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta127` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta128` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta129` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta130` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta131` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta132` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta133` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta134` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta135` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta136` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta137` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta138` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta139` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta140` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta141` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta142` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta143` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta144` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta145` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta146` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta147` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta148` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta149` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta150` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta151` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta152` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta153` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta154` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta155` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta156` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta157` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta158` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta159` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta160` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta161` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta162` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta163` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta164` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta165` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta166` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta167` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta168` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta169` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta170` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta171` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta172` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta173` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta174` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta175` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta176` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta177` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta178` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta179` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta180` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta181` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta182` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta183` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta184` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta185` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta186` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta187` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta188` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta189` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta190` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta191` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta192` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta193` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta194` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta195` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta196` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta197` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta198` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta199` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta200` smallint(6) UNSIGNED NOT NULL DEFAULT 0
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Volcado de datos para la tabla `profesion_secundaria`
--

INSERT INTO `profesion_secundaria` (`user_id`, `profesion`, `receta1`, `receta2`, `receta3`, `receta4`, `receta5`, `receta6`, `receta7`, `receta8`, `receta9`, `receta10`, `receta11`, `receta12`, `receta13`, `receta14`, `receta15`, `receta16`, `receta17`, `receta18`, `receta19`, `receta20`, `receta21`, `receta22`, `receta23`, `receta24`, `receta25`, `receta26`, `receta27`, `receta28`, `receta29`, `receta30`, `receta31`, `receta32`, `receta33`, `receta34`, `receta35`, `receta36`, `receta37`, `receta38`, `receta39`, `receta40`, `receta41`, `receta42`, `receta43`, `receta44`, `receta45`, `receta46`, `receta47`, `receta48`, `receta49`, `receta50`, `receta51`, `receta52`, `receta53`, `receta54`, `receta55`, `receta56`, `receta57`, `receta58`, `receta59`, `receta60`, `receta61`, `receta62`, `receta63`, `receta64`, `receta65`, `receta66`, `receta67`, `receta68`, `receta69`, `receta70`, `receta71`, `receta72`, `receta73`, `receta74`, `receta75`, `receta76`, `receta77`, `receta78`, `receta79`, `receta80`, `receta81`, `receta82`, `receta83`, `receta84`, `receta85`, `receta86`, `receta87`, `receta88`, `receta89`, `receta90`, `receta91`, `receta92`, `receta93`, `receta94`, `receta95`, `receta96`, `receta97`, `receta98`, `receta99`, `receta100`, `receta101`, `receta102`, `receta103`, `receta104`, `receta105`, `receta106`, `receta107`, `receta108`, `receta109`, `receta110`, `receta111`, `receta112`, `receta113`, `receta114`, `receta115`, `receta116`, `receta117`, `receta118`, `receta119`, `receta120`, `receta121`, `receta122`, `receta123`, `receta124`, `receta125`, `receta126`, `receta127`, `receta128`, `receta129`, `receta130`, `receta131`, `receta132`, `receta133`, `receta134`, `receta135`, `receta136`, `receta137`, `receta138`, `receta139`, `receta140`, `receta141`, `receta142`, `receta143`, `receta144`, `receta145`, `receta146`, `receta147`, `receta148`, `receta149`, `receta150`, `receta151`, `receta152`, `receta153`, `receta154`, `receta155`, `receta156`, `receta157`, `receta158`, `receta159`, `receta160`, `receta161`, `receta162`, `receta163`, `receta164`, `receta165`, `receta166`, `receta167`, `receta168`, `receta169`, `receta170`, `receta171`, `receta172`, `receta173`, `receta174`, `receta175`, `receta176`, `receta177`, `receta178`, `receta179`, `receta180`, `receta181`, `receta182`, `receta183`, `receta184`, `receta185`, `receta186`, `receta187`, `receta188`, `receta189`, `receta190`, `receta191`, `receta192`, `receta193`, `receta194`, `receta195`, `receta196`, `receta197`, `receta198`, `receta199`, `receta200`) VALUES
(1, 20, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0),
(2, 20, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `punishment`
--

CREATE TABLE `punishment` (
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `number` tinyint(3) UNSIGNED NOT NULL,
  `reason` varchar(255) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `skillpoint`
--

CREATE TABLE `skillpoint` (
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `sk1` tinyint(3) UNSIGNED NOT NULL,
  `exp1` int(10) UNSIGNED NOT NULL,
  `elu1` int(10) UNSIGNED NOT NULL,
  `sk2` tinyint(3) UNSIGNED NOT NULL,
  `exp2` int(10) UNSIGNED NOT NULL,
  `elu2` int(10) UNSIGNED NOT NULL,
  `sk3` tinyint(3) UNSIGNED NOT NULL,
  `exp3` int(10) UNSIGNED NOT NULL,
  `elu3` int(10) UNSIGNED NOT NULL,
  `sk4` tinyint(3) UNSIGNED NOT NULL,
  `exp4` int(10) UNSIGNED NOT NULL,
  `elu4` int(10) UNSIGNED NOT NULL,
  `sk5` tinyint(3) UNSIGNED NOT NULL,
  `exp5` int(10) UNSIGNED NOT NULL,
  `elu5` int(10) UNSIGNED NOT NULL,
  `sk6` tinyint(3) UNSIGNED NOT NULL,
  `exp6` int(10) UNSIGNED NOT NULL,
  `elu6` int(10) UNSIGNED NOT NULL,
  `sk7` tinyint(3) UNSIGNED NOT NULL,
  `exp7` int(10) UNSIGNED NOT NULL,
  `elu7` int(10) UNSIGNED NOT NULL,
  `sk8` tinyint(3) UNSIGNED NOT NULL,
  `exp8` int(10) UNSIGNED NOT NULL,
  `elu8` int(10) UNSIGNED NOT NULL,
  `sk9` tinyint(3) UNSIGNED NOT NULL,
  `exp9` int(10) UNSIGNED NOT NULL,
  `elu9` int(10) UNSIGNED NOT NULL,
  `sk10` tinyint(3) UNSIGNED NOT NULL,
  `exp10` int(10) UNSIGNED NOT NULL,
  `elu10` int(10) UNSIGNED NOT NULL,
  `sk11` tinyint(3) UNSIGNED NOT NULL,
  `exp11` int(10) UNSIGNED NOT NULL,
  `elu11` int(10) UNSIGNED NOT NULL,
  `sk12` tinyint(3) UNSIGNED NOT NULL,
  `exp12` int(10) UNSIGNED NOT NULL,
  `elu12` int(10) UNSIGNED NOT NULL,
  `sk13` tinyint(3) UNSIGNED NOT NULL,
  `exp13` int(10) UNSIGNED NOT NULL,
  `elu13` int(10) UNSIGNED NOT NULL,
  `sk14` tinyint(3) UNSIGNED NOT NULL,
  `exp14` int(10) UNSIGNED NOT NULL,
  `elu14` int(10) UNSIGNED NOT NULL,
  `sk15` tinyint(3) UNSIGNED NOT NULL,
  `exp15` int(10) UNSIGNED NOT NULL,
  `elu15` int(10) UNSIGNED NOT NULL,
  `sk16` tinyint(3) UNSIGNED NOT NULL,
  `exp16` int(10) UNSIGNED NOT NULL,
  `elu16` int(10) UNSIGNED NOT NULL,
  `sk17` tinyint(3) UNSIGNED NOT NULL,
  `exp17` int(10) UNSIGNED NOT NULL,
  `elu17` int(10) UNSIGNED NOT NULL,
  `sk18` tinyint(3) UNSIGNED NOT NULL,
  `exp18` int(10) UNSIGNED NOT NULL,
  `elu18` int(10) UNSIGNED NOT NULL,
  `sk19` tinyint(3) UNSIGNED NOT NULL,
  `exp19` int(10) UNSIGNED NOT NULL,
  `elu19` int(10) UNSIGNED NOT NULL,
  `sk20` tinyint(3) UNSIGNED NOT NULL,
  `exp20` int(10) UNSIGNED NOT NULL,
  `elu20` int(10) UNSIGNED NOT NULL,
  `sk21` tinyint(3) UNSIGNED NOT NULL,
  `exp21` int(10) UNSIGNED NOT NULL,
  `elu21` int(10) UNSIGNED NOT NULL,
  `sk22` tinyint(3) UNSIGNED NOT NULL,
  `exp22` int(10) UNSIGNED NOT NULL,
  `elu22` int(10) UNSIGNED NOT NULL,
  `sk23` tinyint(3) UNSIGNED NOT NULL,
  `exp23` int(10) UNSIGNED NOT NULL,
  `elu23` int(10) UNSIGNED NOT NULL,
  `sk24` tinyint(3) UNSIGNED NOT NULL,
  `exp24` int(10) UNSIGNED NOT NULL,
  `elu24` int(10) UNSIGNED NOT NULL,
  `sk25` tinyint(3) UNSIGNED NOT NULL,
  `exp25` int(10) UNSIGNED NOT NULL,
  `elu25` int(10) UNSIGNED NOT NULL,
  `sk26` tinyint(3) UNSIGNED NOT NULL,
  `exp26` int(10) UNSIGNED NOT NULL,
  `elu26` int(10) UNSIGNED NOT NULL,
  `sk27` tinyint(3) UNSIGNED NOT NULL,
  `exp27` int(10) UNSIGNED NOT NULL,
  `elu27` int(10) UNSIGNED NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Volcado de datos para la tabla `skillpoint`
--

INSERT INTO `skillpoint` (`user_id`, `sk1`, `exp1`, `elu1`, `sk2`, `exp2`, `elu2`, `sk3`, `exp3`, `elu3`, `sk4`, `exp4`, `elu4`, `sk5`, `exp5`, `elu5`, `sk6`, `exp6`, `elu6`, `sk7`, `exp7`, `elu7`, `sk8`, `exp8`, `elu8`, `sk9`, `exp9`, `elu9`, `sk10`, `exp10`, `elu10`, `sk11`, `exp11`, `elu11`, `sk12`, `exp12`, `elu12`, `sk13`, `exp13`, `elu13`, `sk14`, `exp14`, `elu14`, `sk15`, `exp15`, `elu15`, `sk16`, `exp16`, `elu16`, `sk17`, `exp17`, `elu17`, `sk18`, `exp18`, `elu18`, `sk19`, `exp19`, `elu19`, `sk20`, `exp20`, `elu20`, `sk21`, `exp21`, `elu21`, `sk22`, `exp22`, `elu22`, `sk23`, `exp23`, `elu23`, `sk24`, `exp24`, `elu24`, `sk25`, `exp25`, `elu25`, `sk26`, `exp26`, `elu26`, `sk27`, `exp27`, `elu27`) VALUES
(1, 0, 0, 200, 4, 182, 243, 0, 0, 200, 0, 0, 200, 0, 0, 200, 0, 0, 200, 0, 0, 200, 10, 0, 326, 0, 0, 200, 0, 0, 200, 0, 0, 200, 0, 0, 200, 0, 0, 200, 0, 0, 200, 0, 175, 200, 0, 0, 200, 0, 0, 200, 0, 0, 200, 1, 165, 210, 1, 60, 210, 0, 0, 200, 100, 0, 200, 0, 250, 255, 0, 0, 210, 0, 0, 200, 70, 0, 6085, 100, 0, 200),
(2, 37, 381, 1216, 24, 556, 645, 0, 0, 200, 0, 0, 200, 0, 0, 200, 0, 0, 200, 13, 138, 377, 32, 832, 953, 0, 25, 200, 36, 699, 1158, 0, 0, 200, 0, 0, 200, 0, 0, 200, 0, 0, 200, 40, 742, 1408, 0, 0, 200, 0, 0, 200, 9, 175, 310, 0, 0, 200, 42, 241, 1552, 0, 0, 200, 0, 0, 200, 27, 443, 747, 0, 0, 200, 0, 0, 200, 0, 0, 200, 0, 0, 200);

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `spell`
--

CREATE TABLE `spell` (
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `spell_id1` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id2` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id3` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id4` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id5` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id6` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id7` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id8` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id9` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id10` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id11` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id12` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id13` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id14` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id15` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id16` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id17` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id18` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id19` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id20` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id21` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id22` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id23` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id24` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id25` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id26` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id27` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id28` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id29` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id30` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id31` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id32` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id33` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id34` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id35` smallint(5) UNSIGNED DEFAULT 0
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Volcado de datos para la tabla `spell`
--

INSERT INTO `spell` (`user_id`, `spell_id1`, `spell_id2`, `spell_id3`, `spell_id4`, `spell_id5`, `spell_id6`, `spell_id7`, `spell_id8`, `spell_id9`, `spell_id10`, `spell_id11`, `spell_id12`, `spell_id13`, `spell_id14`, `spell_id15`, `spell_id16`, `spell_id17`, `spell_id18`, `spell_id19`, `spell_id20`, `spell_id21`, `spell_id22`, `spell_id23`, `spell_id24`, `spell_id25`, `spell_id26`, `spell_id27`, `spell_id28`, `spell_id29`, `spell_id30`, `spell_id31`, `spell_id32`, `spell_id33`, `spell_id34`, `spell_id35`) VALUES
(1, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0),
(2, 2, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 3, 6, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);

--
-- Índices para tablas volcadas
--

--
-- Indices de la tabla `atributos`
--
ALTER TABLE `atributos`
  ADD PRIMARY KEY (`user_id`);

--
-- Indices de la tabla `banco_items`
--
ALTER TABLE `banco_items`
  ADD KEY `fk_bank_user` (`user_id`);

--
-- Indices de la tabla `familiar`
--
ALTER TABLE `familiar`
  ADD PRIMARY KEY (`user_id`);

--
-- Indices de la tabla `inventario_items`
--
ALTER TABLE `inventario_items`
  ADD KEY `fk_inventory_user` (`user_id`);

--
-- Indices de la tabla `personaje`
--
ALTER TABLE `personaje`
  ADD PRIMARY KEY (`id`),
  ADD KEY `name` (`name`);

--
-- Indices de la tabla `pet`
--
ALTER TABLE `pet`
  ADD PRIMARY KEY (`user_id`);

--
-- Indices de la tabla `punishment`
--
ALTER TABLE `punishment`
  ADD PRIMARY KEY (`user_id`,`number`);

--
-- Indices de la tabla `skillpoint`
--
ALTER TABLE `skillpoint`
  ADD PRIMARY KEY (`user_id`);

--
-- Indices de la tabla `spell`
--
ALTER TABLE `spell`
  ADD PRIMARY KEY (`user_id`);

--
-- AUTO_INCREMENT de las tablas volcadas
--

--
-- AUTO_INCREMENT de la tabla `personaje`
--
ALTER TABLE `personaje`
  MODIFY `id` mediumint(8) UNSIGNED NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=3;

--
-- Restricciones para tablas volcadas
--

--
-- Filtros para la tabla `atributos`
--
ALTER TABLE `atributos`
  ADD CONSTRAINT `fk_atributos_user` FOREIGN KEY (`user_id`) REFERENCES `personaje` (`id`);

--
-- Filtros para la tabla `banco_items`
--
ALTER TABLE `banco_items`
  ADD CONSTRAINT `fk_bank_user` FOREIGN KEY (`user_id`) REFERENCES `personaje` (`id`);

--
-- Filtros para la tabla `familiar`
--
ALTER TABLE `familiar`
  ADD CONSTRAINT `fk_familiar_user` FOREIGN KEY (`user_id`) REFERENCES `personaje` (`id`);

--
-- Filtros para la tabla `inventario_items`
--
ALTER TABLE `inventario_items`
  ADD CONSTRAINT `fk_inventory_user` FOREIGN KEY (`user_id`) REFERENCES `personaje` (`id`);

--
-- Filtros para la tabla `pet`
--
ALTER TABLE `pet`
  ADD CONSTRAINT `fk_pet_user` FOREIGN KEY (`user_id`) REFERENCES `personaje` (`id`);

--
-- Filtros para la tabla `punishment`
--
ALTER TABLE `punishment`
  ADD CONSTRAINT `fk_punishment_user` FOREIGN KEY (`user_id`) REFERENCES `personaje` (`id`);

--
-- Filtros para la tabla `skillpoint`
--
ALTER TABLE `skillpoint`
  ADD CONSTRAINT `fk_skillpoint_user` FOREIGN KEY (`user_id`) REFERENCES `personaje` (`id`);

--
-- Filtros para la tabla `spell`
--
ALTER TABLE `spell`
  ADD CONSTRAINT `fk_spell_user` FOREIGN KEY (`user_id`) REFERENCES `personaje` (`id`);
COMMIT;

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
