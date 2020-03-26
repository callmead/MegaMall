-- MySQL Administrator dump 1.4
--
-- ------------------------------------------------------
-- Server version	4.1.13a-nt


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8 */;

/*!40014 SET @OLD_UNIQUE_CHECKS=@@UNIQUE_CHECKS, UNIQUE_CHECKS=0 */;
/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO' */;


--
-- Create schema mm_pos
--

CREATE DATABASE /*!32312 IF NOT EXISTS*/ mm_pos;
USE mm_pos;

--
-- Table structure for table `mm_pos`.`customer`
--

DROP TABLE IF EXISTS `customer`;
CREATE TABLE `customer` (
  `Customer_ID` varchar(20) NOT NULL default '',
  `Date` date default NULL,
  `Name` varchar(30) default NULL,
  `CNIC_No` varchar(15) default NULL,
  `Address` varchar(50) default NULL,
  `Occupation` varchar(30) default NULL,
  `Phone_No` varchar(15) default NULL,
  `Mobile_No` varchar(15) default NULL,
  `Other_No` varchar(15) default NULL,
  `Total_Bills_Amount` int(11) default NULL,
  `Total_Due` int(11) default NULL,
  `Remarks` varchar(50) default NULL,
  PRIMARY KEY  (`Customer_ID`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `mm_pos`.`customer`
--

/*!40000 ALTER TABLE `customer` DISABLE KEYS */;
INSERT INTO `customer` (`Customer_ID`,`Date`,`Name`,`CNIC_No`,`Address`,`Occupation`,`Phone_No`,`Mobile_No`,`Other_No`,`Total_Bills_Amount`,`Total_Due`,`Remarks`) VALUES 
 ('-','2007-09-30','General Customer Account','-','-','-','-','-','-',263235,14880,'-');
/*!40000 ALTER TABLE `customer` ENABLE KEYS */;


--
-- Table structure for table `mm_pos`.`customer_account`
--

DROP TABLE IF EXISTS `customer_account`;
CREATE TABLE `customer_account` (
  `TID` varchar(20) NOT NULL default '',
  `Customer_ID` varchar(20) default NULL,
  `Date` date default NULL,
  `Invoice_No` varchar(20) default NULL,
  `Total_Amount` int(11) default NULL,
  `Payment_Mode` varchar(15) default NULL,
  `Amount_Paid` int(11) default NULL,
  `Amount_Due` int(11) default NULL,
  `Remarks` varchar(50) default NULL,
  PRIMARY KEY  (`TID`),
  KEY `fk_Inv1_No` (`Invoice_No`),
  CONSTRAINT `fk_Inv1_No` FOREIGN KEY (`Invoice_No`) REFERENCES `sales` (`Invoice_No`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `mm_pos`.`customer_account`
--

/*!40000 ALTER TABLE `customer_account` DISABLE KEYS */;
INSERT INTO `customer_account` (`TID`,`Customer_ID`,`Date`,`Invoice_No`,`Total_Amount`,`Payment_Mode`,`Amount_Paid`,`Amount_Due`,`Remarks`) VALUES 
 ('T200710120209','-','2007-10-01','IN2007101201955',14200,'Cash',14200,0,'Sale Data'),
 ('T2007111202228','-','2007-11-01','IN2007111202149',22800,'Cash',22800,0,'Sale Data'),
 ('T200711193240','-','2007-01-01','IN20071119322',11100,'Cash',11100,0,'Sale Data'),
 ('T2007112193250','-','2007-11-02','IN2007112193247',400,'Cash',400,380,'Sale Data'),
 ('T2007112202426','-','2007-11-02','IN2007112202359',400,'Cash',400,400,'Sale Data'),
 ('T2007112221858','-','2007-11-02','IN2007112221855',7000,'Cash',6300,6300,'Sale Data'),
 ('T2007121202644','-','2007-12-01','IN200712120266',32400,'Cash',32400,0,'Sale Data'),
 ('T20072119359','-','2007-02-01','IN200721193430',6200,'Cash',6200,0,'Sale Data'),
 ('T20073119528','-','2007-03-01','IN200731195156',12400,'Cash',12400,0,'Sale Data'),
 ('T200741195442','-','2007-04-01','IN200741195344',16225,'Cash',16225,0,'Sale Data'),
 ('T200751195723','-','2007-05-01','IN200751195635',14600,'Cash',14600,0,'Sale Data');
INSERT INTO `customer_account` (`TID`,`Customer_ID`,`Date`,`Invoice_No`,`Total_Amount`,`Payment_Mode`,`Amount_Paid`,`Amount_Due`,`Remarks`) VALUES 
 ('T20076120017','-','2007-06-01','IN200761195949',29200,'Cash',29200,0,'Sale Data'),
 ('T20077120812','-','2007-07-01','IN20077120627',17660,'Cash',17660,0,'Sale Data'),
 ('T200781201032','-','2007-08-01','IN200781201015',28400,'Cash',28400,0,'Sale Data'),
 ('T200791115457','-','2007-09-01','IN200791115423',10900,'Cash',10900,0,'Sale Data'),
 ('T200791173157','-','2007-09-01','IN200791173145',3500,'Cash',3500,0,'Sale Data'),
 ('T200792173534','-','2007-09-02','IN20079217350',1650,'Cash',1650,0,'Sale Data'),
 ('T200792173611','-','2007-09-02','IN200792173558',1300,'Cash',1300,0,'Sale Data'),
 ('T200792173719','-','2007-09-02','IN200792173647',7800,'Cash',0,7800,'Sale Data'),
 ('T20079217402','-','2007-09-02','IN200792173959',200,'Cash',200,0,'Sale Data'),
 ('T20079318352','-','2007-09-03','IN200793183436',7550,'Cash',7550,0,'Sale Data'),
 ('T200793183753','-','2007-09-03','IN200793183733',1750,'Cash',1750,0,'Sale Data'),
 ('T200793183945','-','2007-09-03','IN20079318387',9600,'Cash',9600,0,'Sale Data');
INSERT INTO `customer_account` (`TID`,`Customer_ID`,`Date`,`Invoice_No`,`Total_Amount`,`Payment_Mode`,`Amount_Paid`,`Amount_Due`,`Remarks`) VALUES 
 ('T200794184540','-','2007-09-04','IN200794184519',3800,'Cash',3800,0,'Sale Data'),
 ('T200794185648','-','2007-09-04','IN20079418565',1600,'Cash',1600,0,'Sale Data'),
 ('T200794185731','-','2007-09-04','IN20079418573',600,'Cash',600,0,'Sale Data');
/*!40000 ALTER TABLE `customer_account` ENABLE KEYS */;


--
-- Table structure for table `mm_pos`.`expenditure`
--

DROP TABLE IF EXISTS `expenditure`;
CREATE TABLE `expenditure` (
  `TID` varchar(20) NOT NULL default '',
  `Date` date default NULL,
  `Expense_Type` varchar(20) default NULL,
  `Supplier` varchar(30) default NULL,
  `Payment_Mode` varchar(20) default NULL,
  `Particulars` varchar(30) default NULL,
  `Amount` int(11) default NULL,
  `Remarks` varchar(50) default NULL,
  PRIMARY KEY  (`TID`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `mm_pos`.`expenditure`
--

/*!40000 ALTER TABLE `expenditure` DISABLE KEYS */;
INSERT INTO `expenditure` (`TID`,`Date`,`Expense_Type`,`Supplier`,`Payment_Mode`,`Particulars`,`Amount`,`Remarks`) VALUES 
 ('S2007101202020','2007-10-01','General','-','Cash','-',1200,'-'),
 ('S2007111202239','2007-11-01','General','-','Cash','-',1450,'-'),
 ('S20071119337','2007-01-01','General','-','Cash','-',3500,'-'),
 ('S2007112215220','2007-11-02','General','-','Cash','food',500,'-'),
 ('S2007121202548','2007-12-01','General','-','Cash','-',1800,'-'),
 ('S200721193529','2007-02-01','General','-','Cash','-',1800,'-'),
 ('S200731195218','2007-03-01','General','-','Cash','-',1800,'-'),
 ('S200741195310','2007-04-01','General','-','Cash','-',2000,'-'),
 ('S200751195827','2007-05-01','General','-','Cash','-',1500,'-'),
 ('S200761195927','2007-06-01','General','-','Cash','-',1200,'-'),
 ('S2007712067','2007-07-01','General','-','Cash','-',1500,'-'),
 ('S200781201045','2007-08-01','General','-','Cash','-',1900,'-'),
 ('S20079117177','2007-09-01','General','-','Cash','Food',1500,'-'),
 ('S200792174021','2007-09-02','General','-','Cash','Food',1600,'-'),
 ('S200793183617','2007-09-03','General','-','Cash','Food',2000,'Lunch for 2 Officers Added');
INSERT INTO `expenditure` (`TID`,`Date`,`Expense_Type`,`Supplier`,`Payment_Mode`,`Particulars`,`Amount`,`Remarks`) VALUES 
 ('S200794185749','2007-09-04','General','-','Cash','-',1300,'-');
/*!40000 ALTER TABLE `expenditure` ENABLE KEYS */;


--
-- Table structure for table `mm_pos`.`invoice`
--

DROP TABLE IF EXISTS `invoice`;
CREATE TABLE `invoice` (
  `TID` varchar(20) NOT NULL default '',
  `Invoice_No` varchar(20) default NULL,
  `Product_ID` varchar(20) default NULL,
  `Quantity` int(11) default NULL,
  `Price` int(11) default NULL,
  `Net_Total` int(11) default NULL,
  PRIMARY KEY  (`TID`),
  KEY `fk_Inv_No` (`Invoice_No`),
  KEY `fk_prod_id` (`Product_ID`),
  CONSTRAINT `fk_Inv_No` FOREIGN KEY (`Invoice_No`) REFERENCES `sales` (`Invoice_No`),
  CONSTRAINT `fk_prod_id` FOREIGN KEY (`Product_ID`) REFERENCES `stock` (`Product_ID`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `mm_pos`.`invoice`
--

/*!40000 ALTER TABLE `invoice` DISABLE KEYS */;
INSERT INTO `invoice` (`TID`,`Invoice_No`,`Product_ID`,`Quantity`,`Price`,`Net_Total`) VALUES 
 ('T2007101201955','IN2007101201955','P2007930111625',2,3500,7000),
 ('T200710120209','IN2007101201955','P200793011352',4,1800,7200),
 ('T2007111202149','IN2007111202149','P2007930113011',8,550,4400),
 ('T2007111202152','IN2007111202149','P2007930113234',8,500,4000),
 ('T2007111202228','IN2007111202149','P200793011352',8,1800,14400),
 ('T20071119322','IN20071119322','P2007930112711',2,200,400),
 ('T200711193220','IN20071119322','P2007930111625',1,3500,3500),
 ('T200711193240','IN20071119322','P200793011352',4,1800,7200),
 ('T2007112193250','IN2007112193247','P2007930112711',2,200,400),
 ('T2007112202426','IN2007112202359','P2007930112711',2,200,400),
 ('T2007112221858','IN2007112221855','P2007930111625',2,3500,7000),
 ('T2007121202615','IN200712120266','P2007930113331',2,700,1400),
 ('T2007121202624','IN200712120266','P200793011352',8,1800,14400),
 ('T2007121202644','IN200712120266','P200793011558',8,325,2600),
 ('T200712120266','IN200712120266','P2007930111625',4,3500,14000);
INSERT INTO `invoice` (`TID`,`Invoice_No`,`Product_ID`,`Quantity`,`Price`,`Net_Total`) VALUES 
 ('T200721193430','IN200721193430','P2007930113331',4,700,2800),
 ('T200721193435','IN200721193430','P2007930113234',2,500,1000),
 ('T200721193446','IN200721193430','P200793011558',4,325,1300),
 ('T20072119359','IN200721193430','P2007930113011',2,550,1100),
 ('T200731195156','IN200731195156','P2007930111625',2,3500,7000),
 ('T20073119528','IN200731195156','P200793011352',3,1800,5400),
 ('T200741195344','IN200741195344','P2007930113331',4,700,2800),
 ('T200741195347','IN200741195344','P2007930113011',2,550,1100),
 ('T200741195357','IN200741195344','P2007930111048',10,150,1500),
 ('T200741195418','IN200741195344','P2007930113234',4,500,2000),
 ('T200741195442','IN200741195344','P200793011352',4,1800,7200),
 ('T20074119549','IN200741195344','P200793011558',5,325,1625),
 ('T200751195635','IN200751195635','P2007930111625',2,3500,7000),
 ('T200751195637','IN200751195635','P2007930113331',4,700,2800),
 ('T200751195649','IN200751195635','P2007930113011',4,550,2200);
INSERT INTO `invoice` (`TID`,`Invoice_No`,`Product_ID`,`Quantity`,`Price`,`Net_Total`) VALUES 
 ('T200751195658','IN200751195635','P2007930111456',8,250,2000),
 ('T200751195723','IN200751195635','P2007930111048',4,150,600),
 ('T200761195949','IN200761195949','P2007930112711',4,200,800),
 ('T200761195953','IN200761195949','P2007930111625',4,3500,14000),
 ('T20076120017','IN200761195949','P200793011352',8,1800,14400),
 ('T20077120627','IN20077120627','P2007930112711',2,200,400),
 ('T20077120630','IN20077120627','P2007930111625',2,3500,7000),
 ('T20077120639','IN20077120627','P200793011352',2,1800,3600),
 ('T20077120647','IN20077120627','P2007930113331',1,700,700),
 ('T20077120656','IN20077120627','P2007930113234',2,500,1000),
 ('T20077120714','IN20077120627','P2007930113011',2,550,1100),
 ('T20077120723','IN20077120627','P2007930111456',4,250,1000),
 ('T2007712074','IN20077120627','P200793011558',4,325,1300),
 ('T20077120740','IN20077120627','P2007930111153',2,180,360),
 ('T20077120753','IN20077120627','P2007930111048',4,150,600),
 ('T20077120812','IN20077120627','P2007930113944',6,100,600);
INSERT INTO `invoice` (`TID`,`Invoice_No`,`Product_ID`,`Quantity`,`Price`,`Net_Total`) VALUES 
 ('T200781201015','IN200781201015','P2007930111625',4,3500,14000),
 ('T200781201032','IN200781201015','P200793011352',8,1800,14400),
 ('T200791115423','IN200791115423','P2007930112711',2,200,400),
 ('T200791115443','IN200791115423','P2007930111625',3,3500,10500),
 ('T200791115457','IN200791115423','P200793011352',1,1800,1800),
 ('T200791173145','IN200791173145','P2007930111625',1,3500,3500),
 ('T200791173157','IN200791173145','P2007930113331',1,700,700),
 ('T20079217350','IN20079217350','P2007930113234',2,500,1000),
 ('T200792173534','IN20079217350','P200793011558',2,325,650),
 ('T200792173558','IN200792173558','P2007930111048',5,150,750),
 ('T200792173611','IN200792173558','P2007930113011',1,550,550),
 ('T200792173647','IN200792173647','P2007930113331',1,700,700),
 ('T200792173659','IN200792173647','P200793011352',2,1800,3600),
 ('T200792173719','IN200792173647','P2007930111625',1,3500,3500),
 ('T20079217402','IN200792173959','P2007930112711',1,200,200);
INSERT INTO `invoice` (`TID`,`Invoice_No`,`Product_ID`,`Quantity`,`Price`,`Net_Total`) VALUES 
 ('T200793183436','IN200793183436','P2007930111625',2,3500,7000),
 ('T20079318352','IN200793183436','P2007930113011',1,550,550),
 ('T200793183733','IN200793183733','P2007930111048',3,150,450),
 ('T200793183753','IN200793183733','P200793011558',4,325,1300),
 ('T200793183812','IN20079318387','P2007930113011',1,550,550),
 ('T200793183823','IN20079318387','P2007930113234',2,500,1000),
 ('T200793183834','IN20079318387','P2007930113331',2,700,1400),
 ('T200793183846','IN20079318387','P2007930111625',1,3500,3500),
 ('T200793183859','IN20079318387','P2007930112711',2,200,400),
 ('T20079318387','IN20079318387','P2007930111048',2,150,300),
 ('T200793183911','IN20079318387','P200793011558',2,325,650),
 ('T200793183945','IN20079318387','P200793011352',1,1800,1800),
 ('T200794184519','IN200794184519','P2007930112711',1,200,200),
 ('T200794184540','IN200794184519','P200793011352',2,1800,3600),
 ('T200794185648','IN20079418565','P2007930113944',5,100,500),
 ('T20079418565','IN20079418565','P2007930113011',2,550,1100);
INSERT INTO `invoice` (`TID`,`Invoice_No`,`Product_ID`,`Quantity`,`Price`,`Net_Total`) VALUES 
 ('T200794185731','IN20079418573','P2007930111048',4,150,600);
/*!40000 ALTER TABLE `invoice` ENABLE KEYS */;


--
-- Table structure for table `mm_pos`.`login`
--

DROP TABLE IF EXISTS `login`;
CREATE TABLE `login` (
  `User` varchar(15) default NULL,
  `Password` varchar(10) default NULL,
  `Type` varchar(15) default NULL,
  `Name` varchar(20) default NULL,
  `Designation` varchar(20) default NULL,
  `Remarks` varchar(50) default NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `mm_pos`.`login`
--

/*!40000 ALTER TABLE `login` DISABLE KEYS */;
INSERT INTO `login` (`User`,`Password`,`Type`,`Name`,`Designation`,`Remarks`) VALUES 
 ('admin','admin','Admin','Admin User','Administration','-'),
 ('manager','manager','Manager','Manager User','Management','-'),
 ('salesman','salesman','Salesman','Sales User','Sales Dept','-');
/*!40000 ALTER TABLE `login` ENABLE KEYS */;


--
-- Table structure for table `mm_pos`.`po_details`
--

DROP TABLE IF EXISTS `po_details`;
CREATE TABLE `po_details` (
  `TID` varchar(20) NOT NULL default '',
  `PO_No` varchar(20) default NULL,
  `Product` varchar(20) default NULL,
  `Product_Type` varchar(20) default NULL,
  `Product_Size` varchar(10) default NULL,
  `Quantity` int(11) default NULL,
  `Description` varchar(30) default NULL,
  PRIMARY KEY  (`TID`),
  KEY `fk_PO1_No` (`PO_No`),
  CONSTRAINT `fk_PO1_No` FOREIGN KEY (`PO_No`) REFERENCES `purchase_order` (`PO_No`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `mm_pos`.`po_details`
--

/*!40000 ALTER TABLE `po_details` DISABLE KEYS */;
INSERT INTO `po_details` (`TID`,`PO_No`,`Product`,`Product_Type`,`Product_Size`,`Quantity`,`Description`) VALUES 
 ('T200791121923','PO20079112196','CAPS','CAPS','ALL',50,'-'),
 ('T20079112196','PO20079112196','BLACK SHOES','SHOES','12',20,'Leather'),
 ('T200791122025','PO20079112196','JEANS','JEANS','S,M,L',50,'-');
/*!40000 ALTER TABLE `po_details` ENABLE KEYS */;


--
-- Table structure for table `mm_pos`.`profit`
--

DROP TABLE IF EXISTS `profit`;
CREATE TABLE `profit` (
  `TID` varchar(15) default NULL,
  `Date` date default NULL,
  `year` varchar(5) default NULL,
  `month` varchar(15) default NULL,
  `Expense` int(11) default NULL,
  `Sale` int(11) default NULL,
  `ActualSale` int(11) default NULL,
  `Profit` int(11) default NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `mm_pos`.`profit`
--

/*!40000 ALTER TABLE `profit` DISABLE KEYS */;
INSERT INTO `profit` (`TID`,`Date`,`year`,`month`,`Expense`,`Sale`,`ActualSale`,`Profit`) VALUES 
 ('T200791172210','2007-09-01','2007','09-September',1500,14400,9150,3750),
 ('T200792173736','2007-09-02','2007','09-September',1600,10950,7350,2000),
 ('T200793183553','2007-09-03','2007','09-September',2000,18900,13900,3000),
 ('T200794184555','2007-09-04','2007','09-September',1300,6000,4950,-250),
 ('T200711193331','2007-01-01','2007','01-January',3500,11100,8800,-1200),
 ('T200721193540','2007-02-01','2007','02-February',1800,6200,3300,1200),
 ('T200731193640','2007-03-01','2007','03-March',1800,12400,9500,1100),
 ('T200741195456','2007-04-01','2007','04-April',2000,16225,10975,3250),
 ('T200751195733','2007-05-01','2007','05-May',1500,14600,10800,2300),
 ('T20076120029','2007-06-01','2007','06-June',1200,29200,21400,6600),
 ('T2007712017','2007-07-01','2007','07-July',1500,17660,13490,2670),
 ('T200781201053','2007-08-01','2007','08-August',1900,28400,22000,4500),
 ('T2007101202055','2007-10-01','2007','10-October',1200,14200,11000,2000),
 ('T2007111202246','2007-11-01','2007','11-November',1450,22800,18800,2550);
INSERT INTO `profit` (`TID`,`Date`,`year`,`month`,`Expense`,`Sale`,`ActualSale`,`Profit`) VALUES 
 ('T200712120260','2007-12-01','2007','12-December',1800,32400,25300,5300),
 ('T2007112193351','2007-11-02','2007','11-November',500,7800,5300,1000);
/*!40000 ALTER TABLE `profit` ENABLE KEYS */;


--
-- Table structure for table `mm_pos`.`purchase_order`
--

DROP TABLE IF EXISTS `purchase_order`;
CREATE TABLE `purchase_order` (
  `PO_No` varchar(20) NOT NULL default '',
  `Date` date default NULL,
  `Supplier_ID` varchar(20) default NULL,
  `Delivery_Date` date default NULL,
  `Remarks` varchar(50) default NULL,
  PRIMARY KEY  (`PO_No`),
  KEY `fk_supp_id1` (`Supplier_ID`),
  CONSTRAINT `fk_supp_id1` FOREIGN KEY (`Supplier_ID`) REFERENCES `supplier` (`Supplier_ID`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `mm_pos`.`purchase_order`
--

/*!40000 ALTER TABLE `purchase_order` DISABLE KEYS */;
INSERT INTO `purchase_order` (`PO_No`,`Date`,`Supplier_ID`,`Delivery_Date`,`Remarks`) VALUES 
 ('PO20079112196','2007-09-01','S2007930105724','2007-09-01','-');
/*!40000 ALTER TABLE `purchase_order` ENABLE KEYS */;


--
-- Table structure for table `mm_pos`.`receivings`
--

DROP TABLE IF EXISTS `receivings`;
CREATE TABLE `receivings` (
  `TID` varchar(20) NOT NULL default '',
  `Date` date default NULL,
  `PO_No` varchar(20) default NULL,
  `Product_ID` varchar(20) default NULL,
  `Quantity` int(11) default NULL,
  `Price` int(11) default NULL,
  `Price_per_unit` int(11) default NULL,
  `Remarks` varchar(50) default NULL,
  PRIMARY KEY  (`TID`),
  KEY `fk_PO_No1` (`PO_No`),
  KEY `fk_prod_id3` (`Product_ID`),
  CONSTRAINT `fk_PO_No1` FOREIGN KEY (`PO_No`) REFERENCES `purchase_order` (`PO_No`),
  CONSTRAINT `fk_prod_id3` FOREIGN KEY (`Product_ID`) REFERENCES `stock` (`Product_ID`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `mm_pos`.`receivings`
--

/*!40000 ALTER TABLE `receivings` DISABLE KEYS */;
INSERT INTO `receivings` (`TID`,`Date`,`PO_No`,`Product_ID`,`Quantity`,`Price`,`Price_per_unit`,`Remarks`) VALUES 
 ('T200791164155','2007-09-01','PO20079112196','P2007930113011',20,20000,1000,'-'),
 ('T200791164231','2007-09-01','PO20079112196','P2007930112711',50,15000,300,'-'),
 ('T200791164315','2007-09-01','PO20079112196','P2007930113331',25,25000,1000,'-');
/*!40000 ALTER TABLE `receivings` ENABLE KEYS */;


--
-- Table structure for table `mm_pos`.`sales`
--

DROP TABLE IF EXISTS `sales`;
CREATE TABLE `sales` (
  `Invoice_No` varchar(20) NOT NULL default '',
  `Date` date default NULL,
  `Salesman` varchar(20) default NULL,
  `Customer_ID` varchar(20) default NULL,
  `Grand_Total` int(11) default NULL,
  `Discount` varchar(10) default NULL,
  `Payment_Mode` varchar(15) default NULL,
  `Amount_Paid` int(11) default NULL,
  `Amount_Change` int(11) default NULL,
  `Amount_Due` int(11) default NULL,
  `Buying_Total` int(11) default NULL,
  `Profit` int(11) default NULL,
  `Remarks` varchar(50) default NULL,
  PRIMARY KEY  (`Invoice_No`),
  KEY `fk_cust_id` (`Customer_ID`),
  CONSTRAINT `fk_cust_id` FOREIGN KEY (`Customer_ID`) REFERENCES `customer` (`Customer_ID`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `mm_pos`.`sales`
--

/*!40000 ALTER TABLE `sales` DISABLE KEYS */;
INSERT INTO `sales` (`Invoice_No`,`Date`,`Salesman`,`Customer_ID`,`Grand_Total`,`Discount`,`Payment_Mode`,`Amount_Paid`,`Amount_Change`,`Amount_Due`,`Buying_Total`,`Profit`,`Remarks`) VALUES 
 ('IN2007101201955','2007-10-01','admin','-',14200,'0','Cash',14200,0,0,0,3200,'-'),
 ('IN2007111202149','2007-11-01','admin','-',22800,'0','Cash',22800,0,0,0,4000,'-'),
 ('IN20071119322','2007-01-01','admin','-',11100,'0','Cash',11100,0,0,0,2300,'-'),
 ('IN2007112193247','2007-11-02','admin','-',400,'5%','Cash',400,20,380,0,100,'-'),
 ('IN2007112202359','2007-11-02','admin','-',400,'0%','Cash',400,0,400,300,100,'TEST'),
 ('IN2007112221855','2007-11-02','admin','-',7000,'10%','Cash',6300,0,6300,5000,1300,'-'),
 ('IN200712120266','2007-12-01','admin','-',32400,'0','Cash',32400,0,0,0,7100,'-'),
 ('IN200721193430','2007-02-01','admin','-',6200,'0','Cash',6200,0,0,0,1200,'-'),
 ('IN200731195156','2007-03-01','admin','-',12400,'0','Cash',12400,0,0,0,2900,'-'),
 ('IN200741195344','2007-04-01','admin','-',16225,'0','Cash',16225,0,0,0,3250,'-'),
 ('IN200751195635','2007-05-01','admin','-',14600,'0','Cash',14600,0,0,0,3800,'-');
INSERT INTO `sales` (`Invoice_No`,`Date`,`Salesman`,`Customer_ID`,`Grand_Total`,`Discount`,`Payment_Mode`,`Amount_Paid`,`Amount_Change`,`Amount_Due`,`Buying_Total`,`Profit`,`Remarks`) VALUES 
 ('IN200761195949','2007-06-01','admin','-',29200,'0','Cash',29200,0,0,0,6600,'-'),
 ('IN20077120627','2007-07-01','admin','-',17660,'0','Cash',17660,0,0,0,4170,'-'),
 ('IN200781201015','2007-08-01','admin','-',28400,'0','Cash',28400,0,0,0,6400,'-'),
 ('IN200791115423','2007-09-01','admin','-',10900,'0','Cash',10900,0,0,0,3100,'-'),
 ('IN200791173145','2007-09-01','admin','-',3500,'0','Cash',3500,0,0,0,2150,'-'),
 ('IN20079217350','2007-09-02','admin','-',1650,'0','Cash',1650,0,0,0,400,'-'),
 ('IN200792173558','2007-09-02','admin','-',1300,'0','Cash',1300,0,0,0,700,'-'),
 ('IN200792173647','2007-09-02','admin','-',7800,'0','Cash',7800,0,0,0,2450,'-'),
 ('IN200792173959','2007-09-02','admin','-',200,'0','Cash',200,0,0,0,50,'-'),
 ('IN200793183436','2007-09-03','admin','-',7550,'0','Cash',7550,0,0,0,2050,'-'),
 ('IN200793183733','2007-09-03','admin','-',1750,'0','Cash',1750,0,0,0,350,'-');
INSERT INTO `sales` (`Invoice_No`,`Date`,`Salesman`,`Customer_ID`,`Grand_Total`,`Discount`,`Payment_Mode`,`Amount_Paid`,`Amount_Change`,`Amount_Due`,`Buying_Total`,`Profit`,`Remarks`) VALUES 
 ('IN20079318387','2007-09-03','admin','-',9600,'0','Cash',9600,0,0,0,2600,'-'),
 ('IN200794184519','2007-09-04','admin','-',3800,'0','Cash',3800,0,0,0,650,'-'),
 ('IN20079418565','2007-09-04','admin','-',1600,'0','Cash',1600,0,0,0,100,'-'),
 ('IN20079418573','2007-09-04','admin','-',600,'0','Cash',600,0,0,0,300,'-');
/*!40000 ALTER TABLE `sales` ENABLE KEYS */;


--
-- Table structure for table `mm_pos`.`stock`
--

DROP TABLE IF EXISTS `stock`;
CREATE TABLE `stock` (
  `Product_ID` varchar(20) NOT NULL default '',
  `Date` date default NULL,
  `Product` varchar(30) default NULL,
  `Product_Type` varchar(20) default NULL,
  `Product_Size` varchar(10) default NULL,
  `Company` varchar(30) default NULL,
  `Stock_In_Hand` int(11) default NULL,
  `Description` varchar(25) default NULL,
  `Buying_Price` int(11) default NULL,
  `Selling_Price` int(11) default NULL,
  `ReOrder_Level` int(11) default NULL,
  `Remarks` varchar(50) default NULL,
  PRIMARY KEY  (`Product_ID`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `mm_pos`.`stock`
--

/*!40000 ALTER TABLE `stock` DISABLE KEYS */;
INSERT INTO `stock` (`Product_ID`,`Date`,`Product`,`Product_Type`,`Product_Size`,`Company`,`Stock_In_Hand`,`Description`,`Buying_Price`,`Selling_Price`,`ReOrder_Level`,`Remarks`) VALUES 
 ('P2007930111048','2007-09-30','T-SHIRTS','T-SHIRTS','Small','Bonanza',322,'Mix Colours',100,150,50,'-'),
 ('P2007930111153','2007-09-30','T-SHIRTS','T-SHIRTS','Medium','Bonanza',348,'Mix Colours',120,180,20,'-'),
 ('P2007930111456','2007-09-30','T-SHIRTS','T-SHIRTS','Large','Bonanza',288,'Mic Coulurs',150,250,35,'-'),
 ('P2007930111625','2007-09-30','LEATHER JACKET','JACKETS','All Sizes','ABC',563,'Black',2500,3500,30,'-'),
 ('P2007930112711','2007-09-30','CAPS','CAPS','All Caps','ABC Caps',755,'Mix Colour Caps',150,200,50,'-'),
 ('P2007930113011','2007-09-30','BLACK SHOES','SHOES','Small','Bata',402,'For Childres',500,550,50,'-'),
 ('P2007930113234','2007-09-30','DRESS PANTS','PANTS','28-40','Bonanza',279,'Mix Colour Dress Pants',350,500,20,'-'),
 ('P2007930113331','2007-09-30','JEANS','PANTS','All Sizes','Imported',901,'Mix Colour',550,700,50,'-'),
 ('P200793011352','2007-09-30','RINGS','JEWELRY','All Sizes','-',943,'Silver Rings',1500,1800,100,'-');
INSERT INTO `stock` (`Product_ID`,`Date`,`Product`,`Product_Type`,`Product_Size`,`Company`,`Stock_In_Hand`,`Description`,`Buying_Price`,`Selling_Price`,`ReOrder_Level`,`Remarks`) VALUES 
 ('P2007930113944','2007-09-30','SAMPLE PRODUCT','TEST','Sample','SAMPLE',194,'Sample',100,100,10,'-'),
 ('P200793011558','2007-09-30','DRESS SHIRTS','SHIRTS','14-16','Bonanza',121,'Mix Colour Dress Shirts',275,325,25,'-');
/*!40000 ALTER TABLE `stock` ENABLE KEYS */;


--
-- Table structure for table `mm_pos`.`supplier`
--

DROP TABLE IF EXISTS `supplier`;
CREATE TABLE `supplier` (
  `Supplier_ID` varchar(20) NOT NULL default '',
  `Date` date default NULL,
  `Name` varchar(30) default NULL,
  `Company` varchar(20) default NULL,
  `Contact_Person` varchar(30) default NULL,
  `Address` varchar(50) default NULL,
  `Office_No` varchar(15) default NULL,
  `Mobile_No` varchar(15) default NULL,
  `Other_No` varchar(15) default NULL,
  `Fax_No` varchar(15) default NULL,
  `Total_Bills_Amount` int(11) default NULL,
  `Total_Due` int(11) default NULL,
  `Remarks` varchar(50) default NULL,
  PRIMARY KEY  (`Supplier_ID`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `mm_pos`.`supplier`
--

/*!40000 ALTER TABLE `supplier` DISABLE KEYS */;
INSERT INTO `supplier` (`Supplier_ID`,`Date`,`Name`,`Company`,`Contact_Person`,`Address`,`Office_No`,`Mobile_No`,`Other_No`,`Fax_No`,`Total_Bills_Amount`,`Total_Due`,`Remarks`) VALUES 
 ('S2007930105637','2007-09-30','Hameed','Hameed & Brothers','Munir','Shop 6, Shehzad Market Tarnol, Islamabad','222876541','03225678901','-','-',0,0,'-'),
 ('S2007930105724','2007-09-30','Rashid','Compaq Sollutions','Rashid','IDB Technocare','9895175','01715158718','-','-',60000,10000,'-'),
 ('S2007930105812','2007-09-30','Afzaal Nasim','Noori Traders','Afzaal','E12 5AN East London','228787999','07879261044','-','-',0,0,'-'),
 ('S2007930114014','2007-09-30','Sample Data','Sample','Sample','Sample','000000','000000','000000','000000',0,0,'-');
/*!40000 ALTER TABLE `supplier` ENABLE KEYS */;


--
-- Table structure for table `mm_pos`.`supplier_account`
--

DROP TABLE IF EXISTS `supplier_account`;
CREATE TABLE `supplier_account` (
  `TID` varchar(20) NOT NULL default '',
  `Supplier_ID` varchar(20) default NULL,
  `Date` date default NULL,
  `PO_No` varchar(20) default NULL,
  `Total_Amount` int(11) default NULL,
  `Payment_Mode` varchar(15) default NULL,
  `Paid_Amount` int(11) default NULL,
  `Due_Amount` int(11) default NULL,
  `Remarks` varchar(50) default NULL,
  PRIMARY KEY  (`TID`),
  KEY `fk_supp_id2` (`Supplier_ID`),
  KEY `fk_PO_No` (`PO_No`),
  CONSTRAINT `fk_PO_No` FOREIGN KEY (`PO_No`) REFERENCES `purchase_order` (`PO_No`),
  CONSTRAINT `fk_supp_id2` FOREIGN KEY (`Supplier_ID`) REFERENCES `supplier` (`Supplier_ID`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `mm_pos`.`supplier_account`
--

/*!40000 ALTER TABLE `supplier_account` DISABLE KEYS */;
INSERT INTO `supplier_account` (`TID`,`Supplier_ID`,`Date`,`PO_No`,`Total_Amount`,`Payment_Mode`,`Paid_Amount`,`Due_Amount`,`Remarks`) VALUES 
 ('T200791164353','S2007930105724','2007-09-01','PO20079112196',60000,'Cash',50000,10000,'-');
/*!40000 ALTER TABLE `supplier_account` ENABLE KEYS */;

/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
