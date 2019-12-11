namespace OfficeOpenXml.DataValidation
{
    public enum eDataValidationStorageType
    {
        Unknown,
        Normal,
        X14 //Excel uses this storage type when the data validation contains a reference to another sheet (references through named ranges are excluded)
    }
}
