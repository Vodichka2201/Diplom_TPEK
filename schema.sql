CREATE TABLE `people` (
   `id` bigint unsigned NOT NULL AUTO_INCREMENT,
   `number` bigint unsigned NOT NULL DEFAULT '0',
   `surname` varchar(255) CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci NOT NULL,
   `name` varchar(255) CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci NOT NULL,
   `patronymic` varchar(255) CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci DEFAULT NULL,
   `birth_date` date NOT NULL,
   `gender` enum('мужской','женский') CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci NOT NULL,
   `phone` varchar(255) CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci DEFAULT NULL,
   `is_EPGU` tinyint(1) NOT NULL DEFAULT '0',
   `number_EPGU` varchar(255) CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci DEFAULT NULL,
   `date_EPGU` date DEFAULT NULL,
   `created_at` timestamp NULL DEFAULT NULL,
   `updated_at` timestamp NULL DEFAULT NULL,
   PRIMARY KEY (`id`),
   UNIQUE KEY `unique_person` (`surname`,`name`,`patronymic`,`birth_date`),
   KEY `people_surname_index` (`surname`),
   KEY `people_birth_date_index` (`birth_date`),
   KEY `people_phone_index` (`phone`)
 ) ENGINE=InnoDB AUTO_INCREMENT=189 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
 
 CREATE TABLE `person_certificates` (
   `id` bigint unsigned NOT NULL AUTO_INCREMENT,
   `certificate_type_of_education` varchar(255) CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci DEFAULT NULL,
   `edu_certificate_id` bigint unsigned DEFAULT NULL,
   `certificate_serial_number` varchar(255) CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci DEFAULT NULL,
   `certificate_date` date DEFAULT NULL,
   `average_score` decimal(5,2) DEFAULT NULL,
   `note` text CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci,
   `person_id` bigint unsigned NOT NULL,
   `created_at` timestamp NULL DEFAULT NULL,
   `updated_at` timestamp NULL DEFAULT NULL,
   PRIMARY KEY (`id`),
   KEY `person_certificates_person_id_index` (`person_id`),
   KEY `person_certificates_edu_certificate_id_index` (`edu_certificate_id`),
   KEY `person_certificates_certificate_date_index` (`certificate_date`),
   KEY `person_certificates_average_score_index` (`average_score`),
   CONSTRAINT `person_certificates_edu_certificate_id_foreign` FOREIGN KEY (`edu_certificate_id`) REFERENCES `educational_certificate` (`id`) ON DELETE SET NULL,
   CONSTRAINT `person_certificates_person_id_foreign` FOREIGN KEY (`person_id`) REFERENCES `people` (`id`) ON DELETE CASCADE
 ) ENGINE=InnoDB AUTO_INCREMENT=116 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
 
 CREATE TABLE `educational_certificate` (
   `id` bigint unsigned NOT NULL AUTO_INCREMENT,
   `status_name` varchar(255) CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci DEFAULT NULL,
   `region_id` bigint unsigned DEFAULT NULL,
   `edu_org_full_name` text CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci,
   `edu_org_short_name` text CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci,
   `edu_org_address` text CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci,
   `created_at` timestamp NULL DEFAULT NULL,
   `updated_at` timestamp NULL DEFAULT NULL,
   PRIMARY KEY (`id`),
   KEY `educational_certificate_region_id_index` (`region_id`),
   CONSTRAINT `educational_certificate_region_id_foreign` FOREIGN KEY (`region_id`) REFERENCES `regions` (`id`) ON DELETE SET NULL
 ) ENGINE=InnoDB AUTO_INCREMENT=106653 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
 
 CREATE TABLE `regions` (
   `id` bigint unsigned NOT NULL AUTO_INCREMENT,
   `name` varchar(255) CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci NOT NULL,
   `code` varchar(50) CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci DEFAULT NULL,
   `created_at` timestamp NULL DEFAULT NULL,
   `updated_at` timestamp NULL DEFAULT NULL,
   PRIMARY KEY (`id`),
   UNIQUE KEY `regions_name_unique` (`name`)
 ) ENGINE=InnoDB AUTO_INCREMENT=91 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
 
 CREATE TABLE `educational_levels` (
   `id` bigint unsigned NOT NULL AUTO_INCREMENT,
   `name` varchar(255) CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci NOT NULL,
   `created_at` timestamp NULL DEFAULT NULL,
   `updated_at` timestamp NULL DEFAULT NULL,
   PRIMARY KEY (`id`),
   UNIQUE KEY `educational_levels_name_unique` (`name`)
 ) ENGINE=InnoDB AUTO_INCREMENT=41 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
 
 
CREATE TABLE `certificate_educational_levels` (
   `id` bigint unsigned NOT NULL AUTO_INCREMENT,
   `educational_certificate_id` bigint unsigned NOT NULL,
   `educational_level_id` bigint unsigned NOT NULL,
   `created_at` timestamp NULL DEFAULT NULL,
   `updated_at` timestamp NULL DEFAULT NULL,
   PRIMARY KEY (`id`),
   KEY `idx_edu_cert_id` (`educational_certificate_id`),
   KEY `idx_edu_level_id` (`educational_level_id`)
 ) ENGINE=InnoDB AUTO_INCREMENT=291106 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
 