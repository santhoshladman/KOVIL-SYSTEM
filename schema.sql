-- Kovil Kanakku — MySQL schema
-- Import this file once in phpMyAdmin (or `mysql -u root < schema.sql`) before first use.

CREATE DATABASE IF NOT EXISTS `kovil-backend`
  CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;

USE `kovil-backend`;

CREATE TABLE IF NOT EXISTS app_state (
  id         INT UNSIGNED NOT NULL PRIMARY KEY,
  data       LONGTEXT CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci,
  updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
) ENGINE=InnoDB;

INSERT INTO app_state (id, data)
VALUES (1, NULL)
ON DUPLICATE KEY UPDATE id = id;
