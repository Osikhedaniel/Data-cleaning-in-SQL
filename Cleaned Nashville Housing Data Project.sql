SELECT SaleDate, CONVERT(Date,SaleDate)
FROM Sheet1$_xlnm#_FilterDatabase

UPDATE Emperor.dbo.Sheet1$_xlnm#_FilterDatabase
SET SaleDate = CONVERT(Date,SaleDate)

SELECT SaleDate
FROM Sheet1$_xlnm#_FilterDatabase

ALTER TABLE Sheet1$_xlnm#_FilterDatabase
ADD SaleDateConverted Date;

UPDATE Sheet1$_xlnm#_FilterDatabase
SET SaleDateConverted = CONVERT(Date, SaleDate)

SELECT SaleDateConverted
FROM Sheet1$_xlnm#_FilterDatabase

SELECT PropertyAddress
FROM Sheet1$_xlnm#_FilterDatabase

SELECT A.ParcelID,A.PropertyAddress,B.ParcelID,B.PropertyAddress, ISNULL(A.PropertyAddress,B.PropertyAddress)
FROM Emperor.dbo.Sheet1$_xlnm#_FilterDatabase A
JOIN Emperor.dbo.Sheet1$_xlnm#_FilterDatabase B
   ON A.ParcelID=B.ParcelID
   AND A.[UniqueID ]<>B.[UniqueID ]
   WHERE A.PropertyAddress is null

update A
set PropertyAddress = ISNULL(A.PropertyAddress,B.PropertyAddress)
FROM Emperor.dbo.Sheet1$_xlnm#_FilterDatabase A
JOIN Emperor.dbo.Sheet1$_xlnm#_FilterDatabase B
   ON A.ParcelID=B.ParcelID
   AND A.[UniqueID ]<>B.[UniqueID ]
    WHERE A.PropertyAddress is null

select 
substring(PropertyAddress,1,CHARINDEX(',',PropertyAddress)-1) as Address,
substring(PropertyAddress,CHARINDEX(',',PropertyAddress)+1,LEN(PropertyAddress)) as Address
from Sheet1$_xlnm#_FilterDatabase

alter table Sheet1$_xlnm#_FilterDatabase
add PropertySplitAddress nvarchar(255);

update Sheet1$_xlnm#_FilterDatabase
set PropertySplitAddress = substring(PropertyAddress,1,CHARINDEX(',',PropertyAddress)-1)

alter table Sheet1$_xlnm#_FilterDatabase
add PropertySplitCity nvarchar(255);

update Sheet1$_xlnm#_FilterDatabase
set PropertySplitCity = substring(PropertyAddress,CHARINDEX(',',PropertyAddress)+1,LEN(PropertyAddress)) 

select *
from Sheet1$_xlnm#_FilterDatabase
where LandUse is null

select 
parsename(replace(owneraddress,',','.'),3),
parsename(replace(owneraddress,',','.'),2),
parsename(replace(owneraddress,',','.'),1)
from Sheet1$_xlnm#_FilterDatabase

alter table Sheet1$_xlnm#_FilterDatabase
add CurrentOwnerSplitAddress nvarchar(255);

update Sheet1$_xlnm#_FilterDatabase
set CurrentOwnerSplitAddress = parsename(replace(owneraddress,',','.'),3)

alter table Sheet1$_xlnm#_FilterDatabase
add CurrentOwnerSplitCity nvarchar(255);

update Sheet1$_xlnm#_FilterDatabase
set CurrentOwnerSplitCity = parsename(replace(owneraddress,',','.'),2)

alter table Sheet1$_xlnm#_FilterDatabase
add CurrentOwnerSplitState nvarchar(255);

update Sheet1$_xlnm#_FilterDatabase
set CurrentOwnerSplitState = parsename(replace(owneraddress,',','.'),1)

select *
from Sheet1$_xlnm#_FilterDatabase

select distinct(SoldAsVacant),count(SoldAsVacant)
from Sheet1$_xlnm#_FilterDatabase
group by SoldAsVacant

select SoldAsVacant,
case when SoldAsVacant = 'Y' then 'Yes'
     when SoldAsVacant = 'N' then 'No'
	 else SoldAsVacant
end
from Sheet1$_xlnm#_FilterDatabase

update Sheet1$_xlnm#_FilterDatabase
set SoldAsVacant = case when SoldAsVacant = 'Y' then 'Yes'
     when SoldAsVacant = 'N' then 'No'
	 else SoldAsVacant
end

WITH RowNumCTE AS ( 
select *,
        ROW_NUMBER() over (
		partition by ParcelID,
                     PropertyAddress,
					 SalePrice,
					 Saledate,
					 LegalReference
					 order by
					   uniqueID
					 ) as row_num
from Sheet1$_xlnm#_FilterDatabase)

select *
from RowNumCTE
where row_num > 1
order by PropertyAddress

select *
from Sheet1$_xlnm#_FilterDatabase

alter table Sheet1$_xlnm#_FilterDatabase
drop column OwnerAddress, TaxDistrict, PropertyAddress

alter table Sheet1$_xlnm#_FilterDatabase
drop column SaleDate
