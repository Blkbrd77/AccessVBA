# [Your Database Name] – Documentation for LLM Context

## Database Overview
- **Purpose**: [Write one clear paragraph describing the business problem or process this database supports]
- **Created**: [Month/Year, e.g., March 2022]
- **Last major change**: [Date]
- **File format / backend**: [e.g., .accdb local, split with SQL Server backend, linked tables to SharePoint, etc.]
- **Approximate size**: [~X tables, ~Y queries, ~Z forms/reports, file size in MB]
- **Security / access notes**: [User-level security? Database password? Current user authentication method?]

## Tables
List each table with description, primary key, fields, and important properties.

**Table: [TableName]**  
**Description**: [What this table stores and its role]  
**Primary Key**: [FieldName (Type)]  
**Fields**:
- FieldName1    DataType(Size)    [Notes: PK, Required, Indexed, DefaultValue, ValidationRule, etc.]
- FieldName2    Text(50)          [e.g., Indexed (No Duplicates)]
- CreatedAt     Date/Time         Default = Now()
- Notes         Long Text

**Relationships (outgoing)**:
- 1→Many to [RelatedTable] ([ForeignKey] → [PK], referential integrity: [Yes/No, cascade options])

[Repeat for each important table]

## Relationships Summary
- [TableA] 1→Many [TableB] (enforced, cascade updates)
- [TableB] 1→Many [TableC]
- [Many-to-Many via junction table if present]

## Key Queries
Focus on the most important / frequently used queries.

**qry[NameOfQuery]**  
**Purpose**: [What business question it answers]  
**Type**: [Select / Crosstab / Update / Append / etc.]  
**SQL (simplified or full)**:
```sql
SELECT ...
FROM ...
WHERE ...
