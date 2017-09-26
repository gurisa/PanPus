-- phpMyAdmin SQL Dump
-- version 3.5.0
-- http://www.phpmyadmin.net
--
-- Host: localhost
-- Generation Time: May 21, 2015 at 02:06 PM
-- Server version: 5.0.51
-- PHP Version: 5.2.5

SET SQL_MODE="NO_AUTO_VALUE_ON_ZERO";
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8 */;

--
-- Database: `db_perpus`
--

-- --------------------------------------------------------

--
-- Table structure for table `tb_anggota`
--

CREATE TABLE IF NOT EXISTS `tb_anggota` (
  `id_anggota` int(5) NOT NULL auto_increment,
  `nis_anggota` varchar(250) NOT NULL,
  `nama_anggota` varchar(250) NOT NULL,
  `jenis_kelamin` enum('Pria','Wanita') NOT NULL,
  `kelas_anggota` varchar(250) NOT NULL,
  `jurusan_anggota` varchar(250) NOT NULL,
  `status_anggota` enum('Aktif','Tidak Aktif') NOT NULL,
  `sekolah_anggota` varchar(250) NOT NULL,
  `tanggal_daftar` date NOT NULL,
  `petugas_daftar` varchar(250) NOT NULL,
  `total_denda` int(250) NOT NULL,
  `password_anggota` varchar(250) NOT NULL,
  PRIMARY KEY  (`id_anggota`),
  KEY `nama_anggota` (`nama_anggota`),
  KEY `nama_anggota_2` (`nama_anggota`),
  KEY `nama_anggota_3` (`nama_anggota`),
  KEY `nama_anggota_4` (`nama_anggota`),
  KEY `nama_anggota_5` (`nama_anggota`)
) ENGINE=MyISAM  DEFAULT CHARSET=latin1 AUTO_INCREMENT=32 ;

--
-- Dumping data for table `tb_anggota`
--

INSERT INTO `tb_anggota` (`id_anggota`, `nis_anggota`, `nama_anggota`, `jenis_kelamin`, `kelas_anggota`, `jurusan_anggota`, `status_anggota`, `sekolah_anggota`, `tanggal_daftar`, `petugas_daftar`, `total_denda`, `password_anggota`) VALUES
(12, '1210.276', 'Raka Suryaardi Widjaja', 'Pria', '12', 'Teknik Komputer Dan Jaringan', 'Aktif', 'SMK WIRAKARYA 1 CIPARAY', '2014-10-09', 'admin', 1500, '12111997');

-- --------------------------------------------------------

--
-- Table structure for table `tb_buku`
--

CREATE TABLE IF NOT EXISTS `tb_buku` (
  `id_buku` int(10) NOT NULL auto_increment,
  `nama_buku` varchar(250) NOT NULL,
  `jumlah_buku` int(250) NOT NULL,
  `nama_pengarang` varchar(250) NOT NULL,
  `nama_penerbit` varchar(250) NOT NULL,
  `tahun_terbit` year(4) NOT NULL,
  `tanggal_daftar` date NOT NULL,
  `petugas_daftar` varchar(250) NOT NULL,
  `kategori_buku` varchar(250) NOT NULL,
  PRIMARY KEY  (`id_buku`),
  KEY `nama_buku` (`nama_buku`)
) ENGINE=MyISAM  DEFAULT CHARSET=latin1 AUTO_INCREMENT=11 ;

--
-- Dumping data for table `tb_buku`
--

INSERT INTO `tb_buku` (`id_buku`, `nama_buku`, `jumlah_buku`, `nama_pengarang`, `nama_penerbit`, `tahun_terbit`, `tanggal_daftar`, `petugas_daftar`, `kategori_buku`) VALUES
(1, 'Membuat Aplikasi Perpustakaan Dengan Visual Basic 6 Dan MySQL', 100, 'Raka Suryaardi Widjaja', 'Gurisa', 2014, '2014-10-09', 'admin', 'Pemrograman'),
(3, 'Hari Raya Tionghoa', 100, 'Marcus A.S', 'Suara Harapan Bangsa', 2009, '2015-05-18', 'admin', 'Agama'),
(4, '99,9% Lulus TPA & TKPA SBMPTN', 100, 'Tim Penulis', 'CIF', 2015, '2015-05-18', 'admin', 'SBMPTN'),
(5, 'Pemrograman Aplikasi Android dengan Sencha Touch ', 100, 'Wahana Komputer', 'Andi Publisher', 2015, '2015-05-18', 'admin', 'Pemrograman'),
(6, 'Meningkatkan dan Mencerdaskan IQ Anak Disney 1', 100, 'Disney', 'Elex Media Komputindo ', 2015, '2015-05-18', 'admin', 'Anak'),
(7, 'Laskar Pelangi Song Book', 100, 'Andrea Hirata', 'Bentang Pustaka', 2012, '2015-05-18', 'admin', 'Novel'),
(2, 'Konfigurasi Debian 5 Lenny', 100, 'Pudja Mansyurin', 'AlMansyurin Informatika', 2011, '2015-05-18', 'admin', 'Ebook');

-- --------------------------------------------------------

--
-- Table structure for table `tb_buku_kategori`
--

CREATE TABLE IF NOT EXISTS `tb_buku_kategori` (
  `nama_kategori_buku` varchar(250) NOT NULL,
  PRIMARY KEY  (`nama_kategori_buku`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tb_buku_kategori`
--

INSERT INTO `tb_buku_kategori` (`nama_kategori_buku`) VALUES
('Agama'),
('Anak'),
('Ebook'),
('Novel'),
('Pemrograman'),
('SBMPTN'),
('Sekolah'),
('Umum');

-- --------------------------------------------------------

--
-- Table structure for table `tb_denda`
--

CREATE TABLE IF NOT EXISTS `tb_denda` (
  `id_denda` int(250) NOT NULL auto_increment,
  `id_anggota` int(250) NOT NULL,
  `id_pinjam_detail` int(250) NOT NULL,
  `banyak_denda` int(250) NOT NULL,
  `tanggal_denda` date NOT NULL,
  `petugas_denda` varchar(250) NOT NULL,
  PRIMARY KEY  (`id_denda`),
  KEY `id_pinjam_detail` (`id_pinjam_detail`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1 AUTO_INCREMENT=1 ;

-- --------------------------------------------------------

--
-- Table structure for table `tb_kategori_jurusan`
--

CREATE TABLE IF NOT EXISTS `tb_kategori_jurusan` (
  `nama_kategori_jurusan` varchar(250) NOT NULL,
  PRIMARY KEY  (`nama_kategori_jurusan`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tb_kategori_jurusan`
--

INSERT INTO `tb_kategori_jurusan` (`nama_kategori_jurusan`) VALUES
('Administrasi Perkantoran'),
('Belum Penjurusan'),
('Farmasi Industri'),
('Jasa Boga'),
('Niaga'),
('Pemasaran'),
('Teknik Instalasi Tenaga Listrik'),
('Teknik Komputer Dan Jaringan'),
('Teknik Otomotif'),
('Teknik Permesinan');

-- --------------------------------------------------------

--
-- Table structure for table `tb_kategori_lembaga`
--

CREATE TABLE IF NOT EXISTS `tb_kategori_lembaga` (
  `nama_kategori_lembaga` varchar(250) NOT NULL,
  PRIMARY KEY  (`nama_kategori_lembaga`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tb_kategori_lembaga`
--

INSERT INTO `tb_kategori_lembaga` (`nama_kategori_lembaga`) VALUES
('SDN Majalaya 7'),
('SMK AS-SHIFA CIPARAY'),
('SMK WIRAKARYA 1 CIPARAY'),
('SMK WIRAKARYA 2 CIPARAY');

-- --------------------------------------------------------

--
-- Table structure for table `tb_kategori_tingkat`
--

CREATE TABLE IF NOT EXISTS `tb_kategori_tingkat` (
  `nama_kategori_tingkat` varchar(250) NOT NULL,
  PRIMARY KEY  (`nama_kategori_tingkat`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tb_kategori_tingkat`
--

INSERT INTO `tb_kategori_tingkat` (`nama_kategori_tingkat`) VALUES
('1'),
('10'),
('11'),
('12'),
('2'),
('3'),
('4'),
('5'),
('6'),
('7'),
('8'),
('9');

-- --------------------------------------------------------

--
-- Table structure for table `tb_kembali`
--

CREATE TABLE IF NOT EXISTS `tb_kembali` (
  `id_kembali` int(250) NOT NULL auto_increment,
  `id_pinjam_detail` int(250) NOT NULL,
  `id_anggota` int(250) NOT NULL,
  `nama_petugas` varchar(250) NOT NULL,
  `nama_anggota` varchar(250) NOT NULL,
  `nama_buku` varchar(250) NOT NULL,
  `jumlah_buku` int(250) NOT NULL,
  `tanggal_pinjam` date NOT NULL,
  `tanggal_kembali` date NOT NULL,
  `denda_kembali` int(250) NOT NULL,
  PRIMARY KEY  (`id_kembali`),
  KEY `id_pinjam_detail` (`id_pinjam_detail`),
  KEY `id_anggota` (`id_anggota`)
) ENGINE=MyISAM  DEFAULT CHARSET=latin1 AUTO_INCREMENT=9 ;

-- --------------------------------------------------------

--
-- Table structure for table `tb_petugas`
--

CREATE TABLE IF NOT EXISTS `tb_petugas` (
  `id_petugas` int(3) NOT NULL auto_increment,
  `nama_petugas` varchar(5) NOT NULL,
  `alamat_petugas` text NOT NULL,
  `password_petugas` varchar(250) NOT NULL,
  PRIMARY KEY  (`id_petugas`),
  KEY `nama_petugas` (`nama_petugas`),
  KEY `nama_petugas_2` (`nama_petugas`)
) ENGINE=MyISAM  DEFAULT CHARSET=latin1 AUTO_INCREMENT=227 ;

--
-- Dumping data for table `tb_petugas`
--

INSERT INTO `tb_petugas` (`id_petugas`, `nama_petugas`, `alamat_petugas`, `password_petugas`) VALUES
(1, 'root', 'localheart', 'toor'),
(7, 'admin', 'localserver', 'admin'),
(12, 'raka', 'Jl. Terusan Rancaekek Majalaya No. 289', 'raka');

-- --------------------------------------------------------

--
-- Table structure for table `tb_pinjam`
--

CREATE TABLE IF NOT EXISTS `tb_pinjam` (
  `id_pinjam` int(250) NOT NULL auto_increment,
  `id_anggota` int(250) NOT NULL,
  `nama_anggota` varchar(250) NOT NULL,
  `nama_petugas` varchar(250) NOT NULL,
  `tanggal_pinjam` date NOT NULL,
  `status_pinjam` enum('Pinjam','Kembali') NOT NULL,
  PRIMARY KEY  (`id_pinjam`),
  UNIQUE KEY `id_pinjam_3` (`id_pinjam`),
  KEY `id_pinjam` (`id_pinjam`),
  KEY `id_pinjam_2` (`id_pinjam`),
  KEY `id_anggota` (`id_anggota`),
  KEY `nama_anggota` (`nama_anggota`)
) ENGINE=MyISAM  DEFAULT CHARSET=latin1 AUTO_INCREMENT=10 ;

-- --------------------------------------------------------

--
-- Table structure for table `tb_pinjam_detail`
--

CREATE TABLE IF NOT EXISTS `tb_pinjam_detail` (
  `id_pinjam_detail` int(250) NOT NULL auto_increment,
  `id_pinjam` int(250) NOT NULL,
  `id_buku` int(250) NOT NULL,
  `nama_buku` varchar(250) NOT NULL,
  `jumlah_buku` int(250) NOT NULL,
  `tanggal_pinjam` date NOT NULL,
  `status_pinjam_detail` enum('Pinjam','Kembali') NOT NULL,
  PRIMARY KEY  (`id_pinjam_detail`),
  KEY `id_pinjam` (`id_pinjam`),
  KEY `id_pinjam_2` (`id_pinjam`),
  KEY `id_pinjam_3` (`id_pinjam`),
  KEY `id_pinjam_4` (`id_pinjam`),
  KEY `id_pinjam_5` (`id_pinjam`)
) ENGINE=MyISAM  DEFAULT CHARSET=latin1 AUTO_INCREMENT=13 ;

-- --------------------------------------------------------

--
-- Table structure for table `tb_request`
--

CREATE TABLE IF NOT EXISTS `tb_request` (
  `id_request` int(250) NOT NULL auto_increment,
  `id_pengirim` int(250) NOT NULL,
  `nama_pengirim` varchar(250) NOT NULL,
  `nama_penerima` varchar(250) NOT NULL,
  `id_petugas_tujuan` int(250) NOT NULL,
  `id_anggota_tujuan` int(250) NOT NULL,
  `perihal_request` varchar(250) NOT NULL,
  `konten_request` text NOT NULL,
  `tanggal_request` date NOT NULL,
  `waktu_request` time NOT NULL,
  `status_request` enum('Di Baca','Belum Di Baca') NOT NULL,
  PRIMARY KEY  (`id_request`)
) ENGINE=MyISAM  DEFAULT CHARSET=latin1 AUTO_INCREMENT=26 ;

-- --------------------------------------------------------

--
-- Table structure for table `tb_setting`
--

CREATE TABLE IF NOT EXISTS `tb_setting` (
  `id_setting` int(250) NOT NULL,
  `nama_setting` varchar(250) NOT NULL,
  `status_setting_enum` enum('Ya','Tidak') NOT NULL,
  `status_setting_text` varchar(250) NOT NULL,
  `code_get` varchar(250) NOT NULL,
  `code_send` varchar(250) NOT NULL,
  PRIMARY KEY  (`id_setting`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tb_setting`
--

INSERT INTO `tb_setting` (`id_setting`, `nama_setting`, `status_setting_enum`, `status_setting_text`, `code_get`, `code_send`) VALUES
(1, 'Pop Up About', 'Ya', 'Ya', '', ''),
(2, 'Pop Up Message', 'Ya', 'Ya', '', ''),
(3, 'Aktivasi', 'Tidak', 'Tidak', '', ''),
(4, 'Kelas Anggota', 'Ya', '10', '', ''),
(5, 'Kelas Anggota', 'Ya', '11', '', ''),
(6, 'Kelas Anggota', 'Ya', '12', '', ''),
(7, 'Jurusan Anggota', 'Ya', 'TKJ', '', ''),
(8, 'Jurusan Anggota', 'Ya', 'Otomotif', '', ''),
(9, 'Jurusan Anggota', 'Ya', 'Mesin', '', ''),
(10, 'Jurusan Anggota', 'Ya', 'Listrik', '', ''),
(11, 'Jurusan Anggota', 'Ya', 'Administrasi Perkantoran', '', ''),
(12, 'Jurusan Anggota', 'Ya', 'Niaga', '', ''),
(13, 'Jurusan Anggota', 'Ya', 'Jasa Boga', '', ''),
(14, 'Jurusan Anggota', 'Ya', 'Farmasi', '', ''),
(15, 'Sekolah Anggota', 'Ya', 'SMK WIRAKARYA 1 CIPARAY', '', ''),
(16, 'Sekolah Anggota', 'Ya', 'SMK WIRAKARYA 2 CIPARAY', '', ''),
(17, 'Sekolah Anggota', 'Ya', 'SMK AS-SHIFA CIPARAY', '', ''),
(18, 'Kategori Buku', 'Ya', 'Sekolah', '', ''),
(19, 'Kategori Buku', 'Ya', 'BSE', '', ''),
(20, 'Kategori Buku', 'Ya', 'Novel', '', ''),
(21, 'Kategori Buku', 'Ya', 'Anak - Anak', '', ''),
(22, 'Kategori Buku', 'Ya', 'Remaja', '', ''),
(23, 'Kategori Buku', 'Ya', 'Umum', '', '');

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
