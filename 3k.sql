/*
 Navicat Premium Data Transfer

 Source Server         : 192.168.75.128
 Source Server Type    : MySQL
 Source Server Version : 100610
 Source Host           : 192.168.75.128:3306
 Source Schema         : testdb

 Target Server Type    : MySQL
 Target Server Version : 100610
 File Encoding         : 65001

 Date: 26/10/2023 14:18:57
*/

SET NAMES utf8mb4;
SET FOREIGN_KEY_CHECKS = 0;

-- ----------------------------
-- Table structure for 3k
-- ----------------------------
DROP TABLE IF EXISTS `3k`;
CREATE TABLE `3k`  (
  `unit` varchar(255) CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci NULL DEFAULT NULL COMMENT '所属单位',
  `unit2` varchar(255) CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci NULL DEFAULT NULL COMMENT '所属单位2(和上一个一样,忘了为啥创建这个列了最早的代码丢失,懒得改了)',
  `三` varchar(255) CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci NULL DEFAULT NULL COMMENT '忘了为啥用了',
  `host` varchar(255) CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci NULL DEFAULT NULL COMMENT '主机地址',
  `system` varchar(255) CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci NULL DEFAULT NULL COMMENT '系统名称',
  `CVE编号` varchar(255) CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci NULL DEFAULT NULL COMMENT '漏扫软件输出的漏洞名称,通常最后会带cve编号,程序会自动识别出来,填写到附件1中',
  `危害影响` text CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci NULL COMMENT '危害影响',
  `解决方案` text CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci NULL COMMENT '解决方案'
) ENGINE = InnoDB CHARACTER SET = utf8mb4 COLLATE = utf8mb4_general_ci ROW_FORMAT = Dynamic;

SET FOREIGN_KEY_CHECKS = 1;
