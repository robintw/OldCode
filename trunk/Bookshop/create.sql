CREATE TABLE tblBook (
    ISBN VARCHAR(10) NOT NULL,
    Title VARCHAR(255) NOT NULL,
    Author VARCHAR(255) NOT NULL,
    Publisher VARCHAR(255),
    PRIMARY KEY (ISBN),
    KEY Author(Author),
    KEY Title(Title)
);

CREATE TABLE tblBookCopy (
    CopyID INTEGER NOT NULL,
    ISBN VARCHAR(13) NOT NULL,
    CostPrice FLOAT DEFAULT 0,
    SellingPrice FLOAT NOT NULL DEFAULT 0,
    SecondHand BIT NOT NULL,
    StockCheck BIT NOT NULL,
    Alphacode VARCHAR(6),
    DateBought DATETIME,
    PRIMARY KEY (CopyID),
    KEY Alphacode(Alphacode),
    KEY CopyID(CopyID),
    KEY ISBN(ISBN)
);

CREATE TABLE tblNotFound (
    NotFoundID INTEGER NOT NULL,
    ISBN VARCHAR(50),
    Title VARCHAR(100),
    Author VARCHAR(50),
    Publisher VARCHAR(50),
    PRIMARY KEY (NotFoundID),
    KEY NotFoundID(NotFoundID)
);

CREATE TABLE tblOrder (
    OrderID INTEGER NOT NULL,
    ISBN VARCHAR(50),
    Title VARCHAR(100),
    Author VARCHAR(50),
    Publisher VARCHAR(50),
    DateOrdered DATETIME,
    Alphacode VARCHAR(6),
    PRIMARY KEY (OrderID),
    KEY Author(Author),
    KEY Alphacode(Alphacode),
    KEY OrderID(OrderID),
    KEY Title(Title)
);

CREATE TABLE tblStudent (
    Alphacode VARCHAR(6) NOT NULL,
    FirstName VARCHAR(25) NOT NULL,
    LastName VARCHAR(25) NOT NULL,
    SchoolYear SMALLINT DEFAULT Null,
    House VARCHAR(50),
    Allowance FLOAT DEFAULT Null,
    Spent FLOAT DEFAULT 0,
    Old BIT NOT NULL,
    PRIMARY KEY (Alphacode),
    KEY Alphacode(Alphacode),
    KEY FirstName(FirstName),
    KEY LastName(LastName)
);

CREATE TABLE tblSupplier (
    SupplierAbbrev VARCHAR(3) NOT NULL,
    Name VARCHAR(50) NOT NULL,
    ContactNumber VARCHAR(15),
    PRIMARY KEY (SupplierAbbrev),
    KEY Name(Name)
);

CREATE TABLE tblHouse (
    HouseAbbrev VARCHAR(3) NOT NULL,
    HouseName VARCHAR(15),
    PRIMARY KEY (HouseAbbrev),
    KEY HouseID(HouseAbbrev)
);