-- phpMyAdmin SQL Dump
-- version 5.2.1
-- https://www.phpmyadmin.net/
--
-- Anamakine: 127.0.0.1
-- Üretim Zamanı: 14 Oca 2025, 18:33:09
-- Sunucu sürümü: 10.4.32-MariaDB
-- PHP Sürümü: 8.2.12

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
START TRANSACTION;
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;

--
-- Veritabanı: `chatbot`
--

-- --------------------------------------------------------

--
-- Tablo için tablo yapısı `questions`
--

CREATE TABLE `questions` (
  `soru_ID` int(11) NOT NULL,
  `ogrenci_ID` int(11) NOT NULL,
  `questions` text NOT NULL,
  `answers` text NOT NULL,
  `Feedback` text NOT NULL,
  `topic` varchar(50) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_turkish_ci;

--
-- Tablo döküm verisi `questions`
--

INSERT INTO `questions` (`soru_ID`, `ogrenci_ID`, `questions`, `answers`, `Feedback`, `topic`) VALUES
(1, 12345, 'Senanın doğum tarihi nedir?', '22.05.2003', 'Olumlu', 'Tarih'),
(2, 212523011, 'Atatürk\'ün doğum yılı nedir?', '1881', 'Olumlu', 'Tarih'),
(3, 212523018, 'Koronavirüs kaç yılında Türkiye\'ye gelmiştir', '2020', 'Olumsuz', 'Genel Kültür'),
(4, 212523051, 'Hatay\'ın kurtuluş tarihi nedir?', '23 Temmuz 1939', 'Olumlu', 'Tarih'),
(5, 212523051, 'Adana\'nın kurtuluş tarihi nedir?', '5 Ocak 1922', 'Olumlu', 'Tarih'),
(8, 212523051, 'nasılsın', 'hastayım', 'Olumsuz', 'günlük'),
(9, 212523051, 'nasılsın', 'felfenaa', 'Olumlu', 'günlük'),
(10, 212523051, 'nasılsın', 'yaav', 'Olumlu', 'günlük'),
(11, 212523051, 'güzel misin', 'e herhalde', 'Olumlu', 'günlük'),
(12, 212523051, 'en sevdiğin renk', 'siyah', 'Olumlu', 'günlük'),
(13, 212523018, 'En sevdiğin yemek nedir?', 'Makarna', 'Olumlu', 'günlük'),
(14, 212523026, 'Çiğdem Makine Öğrenmesinden kaç aldı?', '90', 'Olumlu', 'Bilim'),
(15, 212523051, 'Hangi kahveyi seversin?', 'Mocha', 'Olumlu', 'Bilim');

-- --------------------------------------------------------

--
-- Tablo için tablo yapısı `users`
--

CREATE TABLE `users` (
  `id` int(11) NOT NULL,
  `Student_number` int(11) NOT NULL,
  `Password` int(11) NOT NULL,
  `name` varchar(50) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_turkish_ci;

--
-- Tablo döküm verisi `users`
--

INSERT INTO `users` (`id`, `Student_number`, `Password`, `name`) VALUES
(1, 123, 1234, 'deneme'),
(2, 123, 1234, 'deneme'),
(3, 12345, 12345, 'Test'),
(4, 212523051, 654321, 'Ezgi Günay'),
(5, 212523011, 121212, 'Cihan Bulut'),
(6, 212523018, 789, 'Sena Başer'),
(7, 212523026, 123654, 'Çiğdem Pehlivan');

--
-- Dökümü yapılmış tablolar için indeksler
--

--
-- Tablo için indeksler `questions`
--
ALTER TABLE `questions`
  ADD PRIMARY KEY (`soru_ID`);

--
-- Tablo için indeksler `users`
--
ALTER TABLE `users`
  ADD PRIMARY KEY (`id`);

--
-- Dökümü yapılmış tablolar için AUTO_INCREMENT değeri
--

--
-- Tablo için AUTO_INCREMENT değeri `questions`
--
ALTER TABLE `questions`
  MODIFY `soru_ID` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=16;

--
-- Tablo için AUTO_INCREMENT değeri `users`
--
ALTER TABLE `users`
  MODIFY `id` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=8;
COMMIT;

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
