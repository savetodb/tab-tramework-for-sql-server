-- =============================================
-- Tab Framework for Microsoft SQL Server
-- Version 10.8, January 9, 2023
--
-- Copyright 2021-2023 Gartle LLC
--
-- License: MIT
-- =============================================

GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     10.1, 2022-08-17
-- Description: The view generates application handlers
-- =============================================

ALTER VIEW [tab].[xl_app_handlers]
AS

SELECT
    CAST(s.TABLE_SCHEMA AS nvarchar(20)) AS TABLE_SCHEMA
    , CAST(s.TABLE_NAME AS nvarchar(128)) AS TABLE_NAME
    , CAST(s.COLUMN_NAME AS nvarchar(128)) AS COLUMN_NAME
    , CAST(s.EVENT_NAME AS nvarchar(25)) AS EVENT_NAME
    , CAST(s.HANDLER_SCHEMA AS nvarchar(20)) AS HANDLER_SCHEMA
    , CAST(s.HANDLER_NAME AS nvarchar(128)) AS HANDLER_NAME
    , CAST(s.HANDLER_TYPE AS nvarchar(25)) AS HANDLER_TYPE
    , CAST(s.HANDLER_CODE AS nvarchar(max)) HANDLER_CODE
    , CAST(s.TARGET_WORKSHEET AS nvarchar(128)) AS TARGET_WORKSHEET
    , CAST(s.MENU_ORDER AS int) AS MENU_ORDER
    , CAST(s.EDIT_PARAMETERS AS bit) AS EDIT_PARAMETERS
FROM
    (VALUES
        ('tab', 'columns', 'table_id', 'ValidationList', 'tab', 'table_id', 'CODE', 'SELECT id, table_schema + ''.'' + table_name AS name FROM tab.tables ORDER BY table_schema, table_name', NULL, NULL, NULL)
        , ('tab', 'columns', 'type_id', 'ValidationList', 'tab', 'type_id', 'CODE', 'SELECT id, name FROM tab.types WHERE datatype IS NOT NULL ORDER BY name', NULL, NULL, NULL)
        , ('tab', 'columns', 'value_table_id', 'ValidationList', 'tab', 'table_id', 'CODE', 'SELECT id, table_schema + ''.'' + table_name AS name FROM tab.tables ORDER BY table_schema, table_name', NULL, NULL, NULL)
        , ('tab', 'columns', 'value_column_id', 'ValidationList', 'tab', 'value_column_id', 'CODE', 'SELECT c.id, table_schema + ''.'' + table_name + ''.'' + c.name AS name, c.table_id AS value_table_id FROM tab.columns c INNER JOIN tab.tables t ON t.id = c.table_id ORDER BY c.value_table_id, c.column_id, c.name', NULL, NULL, NULL)
        , ('tab', 'columns', 'parameter_name', 'ValidationList', 'tab', 'parameter_name', 'VALUES', ', p1, p2, p3, p4, p5, p6', NULL, NULL, NULL)
        , ('tab', 'usp_select_table', 'table_id', 'ParameterValues', 'tab', 'xl_list_table_id', 'PROCEDURE', NULL, '_NotNull', NULL, NULL)
        , ('tab', 'usp_select_table', NULL, 'JsonForm', NULL, NULL, 'ATTRIBUTE', NULL, NULL, NULL, NULL)
        , ('tab', 'usp_select_table', 'cell_formulas', 'KeepFormulas', NULL, NULL, 'ATTRIBUTE', NULL, NULL, NULL, NULL)
        , ('tab', 'usp_select_table', 'cell_comments', 'KeepComments', NULL, NULL, 'ATTRIBUTE', NULL, NULL, NULL, NULL)
        , ('tab', 'usp_select_translations', NULL, 'JsonForm', NULL, NULL, 'ATTRIBUTE', NULL, NULL, NULL, NULL)
        , ('tab', 'usp_select_translations', NULL, 'ProtectRows', NULL, NULL, 'ATTRIBUTE', NULL, NULL, NULL, NULL)
        , ('tab', 'rows', NULL, 'DoNotAddValidation', NULL, NULL, 'ATTRIBUTE', NULL, NULL, NULL, NULL)
        , ('tab', 'cells', NULL, 'DoNotAddValidation', NULL, NULL, 'ATTRIBUTE', NULL, NULL, NULL, NULL)
        , ('tab', 'cells', NULL, 'DoNotAddManyToMany', NULL, NULL, 'ATTRIBUTE', NULL, NULL, NULL, NULL)
        , ('tab', 'translations', NULL, 'DoNotAddValidation', NULL, NULL, 'ATTRIBUTE', NULL, NULL, NULL, NULL)
        , ('tab', 'tables', NULL, 'License', NULL, NULL, 'ATTRIBUTE', 'QFsOXa3retQAOu4EdB4JRfWDExZ3MovBq1ytjtSxn6gJAGG480B86wzldn8VqJA3icmRZ2x1hlNdyn1F5Xeu1B88jXwnVXHrJr8URaSciVB5OzscmIongev0muuy5eOuR4hFWwb1HAKNJfnGoFxLYNsU79yLRlTiXVPcdTEiPN4=', NULL, NULL, NULL)
        , ('tab', 'columns', NULL, 'License', NULL, NULL, 'ATTRIBUTE', 'ga6j2jOOyR5HtB01NSRWf5bSOMC9qkxpuFmY/X4k8V9b2gqY7SEdFVYPKHVZht2Lo7p5GRFUgiEA5hVtF3kM4wW4S1wghHqrwfxCEcTf7095h0SUAwjBz7WVwbBmHNOSFeib/JDGo3MIxiPehoCEFPD7UMXrzR5TNou4y9x225I=', NULL, NULL, NULL)
        , ('tab', 'languages', NULL, 'License', NULL, NULL, 'ATTRIBUTE', 'xPheHHnyU0PI4CXdf9fIf2T+APWWtbp0Txsz95etKhS1cFSeWEGepMw4fdwEew4y7vs7/YXWWQbVJqc4uFFDxHoSQVhzNajYCLopMYsd85hJHLRVAlvrgwpQUUMKGv/fC2moQnlGNDvrsy+yuOEvYI94oB9IzJyJkqqMsDN4bUo=', NULL, NULL, NULL)
        , ('tab', 'usp_select_table', NULL, 'License', NULL, NULL, 'ATTRIBUTE', 'QXT0BUCJn40gUHKwL4XO2pBvL4V/QjsX0tQJTiqQA3djoIjYn4PambHY9EfjJpb9uAHZOVt170xdkkuhk9Yoj+GfD9N/mrXvJwL2bDfNl/t1nH2uFFO6hvm8zD+hyvh/EHfeUxtXkekT1Z5c8HHVuRN9IAQFtASKc6VCFbcAr7g=', NULL, NULL, NULL)
        , ('tab', 'usp_select_translations', NULL, 'License', NULL, NULL, 'ATTRIBUTE', 'uyjegG3w2S2P1ypzDceg+euiBOYnslyEAl0I2rWOmoi5wQ5R49rnPrmIEs6c2BtXktMBrL0cM/3GDOMVKTokS7QkUglCkvAEnt2vUqUM98jrhbQMicGbFA+ujX14QURnNULveGV116MUDDdvAKm6eV2YUo8uYGx8+eL2Ro10CwM=', NULL, NULL, NULL)
        , ('tab', 'users', NULL, 'ContextMenu', NULL, 'Add {user} to tab_users', 'CODE', 'ALTER ROLE tab_users ADD MEMBER @user', NULL, 1, NULL)
        , ('tab', 'users', NULL, 'ContextMenu', NULL, 'Add {user} to tab_developers', 'CODE', 'ALTER ROLE tab_developers ADD MEMBER @user', NULL, 2, NULL)
        , ('tab', 'users', NULL, 'ContextMenu', NULL, NULL, 'MENUSEPARATOR', NULL, NULL, 3, NULL)
        , ('tab', 'users', NULL, 'ContextMenu', NULL, 'Remove {user} from tab_users', 'CODE', 'ALTER ROLE tab_users DROP MEMBER @user', NULL, 4, NULL)
        , ('tab', 'users', NULL, 'ContextMenu', NULL, 'Remove {user} from tab_developers', 'CODE', 'ALTER ROLE tab_developers DROP MEMBER @user', NULL, 5, NULL)
    ) s(TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS)

UNION ALL
SELECT
    CAST(t.table_schema AS nvarchar(20)) AS TABLE_SCHEMA
    , CAST(t.table_name AS nvarchar(128)) AS TABLE_NAME
    , CAST(s.COLUMN_NAME AS nvarchar(128)) AS COLUMN_NAME
    , CAST(s.EVENT_NAME AS nvarchar(25)) AS EVENT_NAME
    , CAST(s.HANDLER_SCHEMA AS nvarchar(20)) AS HANDLER_SCHEMA
    , CAST(s.HANDLER_NAME AS nvarchar(128)) AS HANDLER_NAME
    , CAST(s.HANDLER_TYPE AS nvarchar(25)) AS HANDLER_TYPE
    , CAST(s.HANDLER_CODE AS nvarchar(max)) HANDLER_CODE
    , CAST(NULL AS nvarchar(128)) AS TARGET_WORKSHEET
    , CAST(NULL AS int) AS MENU_ORDER
    , CAST(NULL AS bit) AS EDIT_PARAMETERS
FROM
    tab.tables t
    CROSS JOIN (VALUES
        (NULL, 'JsonForm', NULL, NULL, 'ATTRIBUTE', NULL, NULL, NULL, NULL)
        , ('cell_formulas', 'KeepFormulas', NULL, NULL, 'ATTRIBUTE', NULL, NULL, NULL, NULL)
        , ('cell_comments', 'KeepComments', NULL, NULL, 'ATTRIBUTE', NULL, NULL, NULL, NULL)
        , (NULL, 'LoadFormat', 'tab', 'formats', 'TABLE', NULL, NULL, NULL, NULL)
    ) s(COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS)
WHERE
    t.store_formulas = 1 OR NOT s.EVENT_NAME IN ('KeepFormulas', 'KeepComments')

UNION ALL
SELECT
    CAST(t.table_schema AS nvarchar(20)) AS TABLE_SCHEMA
    , CAST(t.table_name AS nvarchar(128)) AS TABLE_NAME
    , CAST(s.COLUMN_NAME AS nvarchar(128)) AS COLUMN_NAME
    , CAST(s.EVENT_NAME AS nvarchar(25)) AS EVENT_NAME
    , CAST(s.HANDLER_SCHEMA AS nvarchar(20)) AS HANDLER_SCHEMA
    , CAST(s.HANDLER_NAME AS nvarchar(128)) AS HANDLER_NAME
    , CAST(s.HANDLER_TYPE AS nvarchar(25)) AS HANDLER_TYPE
    , CAST(s.HANDLER_CODE AS nvarchar(max)) HANDLER_CODE
    , CAST(NULL AS nvarchar(128)) AS TARGET_WORKSHEET
    , CAST(NULL AS int) AS MENU_ORDER
    , CAST(NULL AS bit) AS EDIT_PARAMETERS
FROM
    tab.tables t
    CROSS JOIN (VALUES
        (NULL, 'ProtectRows', NULL, NULL, 'ATTRIBUTE', NULL, NULL, NULL, NULL)
        , (NULL, 'DoNotSave', NULL, NULL, 'ATTRIBUTE', NULL, NULL, NULL, NULL)
    ) s(COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS)
WHERE
    (t.protect_rows = 1 AND s.EVENT_NAME IN ('ProtectRows')
    OR t.do_not_save = 1 AND s.EVENT_NAME IN ('DoNotSave')
    ) AND EXISTS (
            SELECT
                p.name
            FROM
                sys.database_principals p
                INNER JOIN sys.database_role_members rm ON rm.member_principal_id = p.principal_id
            WHERE
                p.name = USER_NAME()
                AND rm.role_principal_id = USER_ID('tab_users')
                AND NOT rm.role_principal_id = USER_ID('tab_developers')
            )

UNION ALL
SELECT
    CAST(t.table_schema AS nvarchar(20)) AS TABLE_SCHEMA
    , CAST(t.table_name AS nvarchar(128)) AS TABLE_NAME
    , CAST(c.name AS nvarchar(128)) AS COLUMN_NAME
    , CAST('ValidationList' AS nvarchar(25)) AS EVENT_NAME
    , vt.table_schema AS HANDLER_SCHEMA
    , 'list_' + vt.table_name AS HANDLER_NAME
    , 'CODE' AS HANDLER_TYPE
    , 'EXEC tab.xl_list_by_column_id @column_id=' + CAST(vc.id AS nvarchar) + ', @data_language= @data_language' AS HANDLER_CODE
    , CAST(NULL AS nvarchar(128)) AS TARGET_WORKSHEET
    , CAST(NULL AS int) AS MENU_ORDER
    , CAST(NULL AS bit) AS EDIT_PARAMETERS
FROM
    tab.columns c
    INNER JOIN tab.tables t ON t.id = c.table_id
    INNER JOIN tab.columns vc ON vc.id = c.value_column_id
    INNER JOIN tab.tables vt ON vt.id = vc.table_id

UNION ALL
SELECT
    CAST(t.table_schema AS nvarchar(20)) AS TABLE_SCHEMA
    , CAST(t.table_name AS nvarchar(128)) AS TABLE_NAME
    , CAST(c.name AS nvarchar(128)) AS COLUMN_NAME
    , CAST(s.EVENT_NAME AS nvarchar(25)) AS EVENT_NAME
    , CAST(s.HANDLER_SCHEMA AS nvarchar(20)) AS HANDLER_SCHEMA
    , CAST(s.HANDLER_NAME AS nvarchar(128)) AS HANDLER_NAME
    , CAST(s.HANDLER_TYPE AS nvarchar(25)) AS HANDLER_TYPE
    , CAST(s.HANDLER_CODE AS nvarchar(max)) HANDLER_CODE
    , CAST(NULL AS nvarchar(128)) AS TARGET_WORKSHEET
    , CAST(NULL AS int) AS MENU_ORDER
    , CAST(NULL AS bit) AS EDIT_PARAMETERS
FROM
    tab.columns c
    INNER JOIN tab.tables t ON t.id = c.table_id
    CROSS JOIN (VALUES
        ('DoNotChange', NULL, NULL, 'ATTRIBUTE', NULL, NULL, NULL, NULL)
    ) s(EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS)
WHERE
    (c.do_not_change = 1 AND s.EVENT_NAME IN ('DoNotChange')
    ) AND EXISTS (
            SELECT
                p.name
            FROM
                sys.database_principals p
                INNER JOIN sys.database_role_members rm ON rm.member_principal_id = p.principal_id
            WHERE
                p.name = USER_NAME()
                AND rm.role_principal_id = USER_ID('tab_users')
                AND NOT rm.role_principal_id = USER_ID('tab_developers')
            )

UNION ALL
SELECT
    CAST(t.table_schema AS nvarchar(20)) AS TABLE_SCHEMA
    , CAST(t.table_name AS nvarchar(128)) AS TABLE_NAME
    , CAST(c.name AS nvarchar(128)) AS COLUMN_NAME
    , CAST(ct.datatype AS nvarchar(25)) AS EVENT_NAME
    , CAST(NULL AS nvarchar(20)) AS HANDLER_SCHEMA
    , CAST(NULL AS nvarchar(128)) AS HANDLER_NAME
    , CAST('ATTRIBUTE' AS nvarchar(25)) AS HANDLER_TYPE
    , CAST(NULL AS nvarchar(max)) HANDLER_CODE
    , CAST(NULL AS nvarchar(128)) AS TARGET_WORKSHEET
    , CAST(NULL AS int) AS MENU_ORDER
    , CAST(NULL AS bit) AS EDIT_PARAMETERS
FROM
    tab.columns c
    INNER JOIN tab.tables t ON t.id = c.table_id
    INNER JOIN tab.types ct ON ct.id = c.type_id
WHERE
    ct.datatype IS NOT NULL

UNION ALL
SELECT
    CAST(t.table_schema AS nvarchar(20)) AS TABLE_SCHEMA
    , CAST(t.table_name AS nvarchar(128)) AS TABLE_NAME
    , CAST(c.parameter_name AS nvarchar(128)) AS COLUMN_NAME
    , CAST(ct.datatype AS nvarchar(25)) AS EVENT_NAME
    , CAST(NULL AS nvarchar(20)) AS HANDLER_SCHEMA
    , CAST(NULL AS nvarchar(128)) AS HANDLER_NAME
    , CAST('ATTRIBUTE' AS nvarchar(25)) AS HANDLER_TYPE
    , CAST(NULL AS nvarchar(max)) HANDLER_CODE
    , CAST(NULL AS nvarchar(128)) AS TARGET_WORKSHEET
    , CAST(NULL AS int) AS MENU_ORDER
    , CAST(NULL AS bit) AS EDIT_PARAMETERS
FROM
    tab.columns c
    INNER JOIN tab.tables t ON t.id = c.table_id
    INNER JOIN tab.types ct ON ct.id = c.type_id
WHERE
    ct.datatype IS NOT NULL
    AND c.parameter_name IS NOT NULL


GO

print 'Tab Framework updated';
