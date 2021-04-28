import pyodbc

# Keep an array of column names and another array of column data
# have get(column_name) and set(column_name, value) methods
# for get and set, alias Export Info as "id"
# class Tag:
#     _column_names = {}

#     def __init__(self, tag_type = None, values = None):
#         self.tag_type = tag_type
#         self.column_names = [] if values == None else values
#         self.values = [] if values == None else values

#     # Not yet implemented
#     @staticmethod
#     def assumed_type_ab(name):
#         # TODO: Should allow user to load this in from a file (same for column_names?)
#         tagTypes =  [    
#             ("AB_AI", [ "LT", "LIT", "AIT", "FIT", "PIT", "TT", "WIT", "TT", "ZA", "ZS" ] ),
#             ("AB_DA", [ "TAH", "TAL", "LAHH", "SD", "LS", "VAH", "PAH", "PAL", "FAL", "LAL", "LAH" ] ),
#             ("AB_FV", [ "FV" ] ),
#             ("AB_FCV", [ "FCV" ] ),
#             ("AB_MOTOR", [ "P", "CP", "SC", "BL", "CF" ] ),
#             ("AB_TOTALIZER", [ ] )]
#         name = name.split('\\\\')[-1]
#         name = name.split('_')[0]
#         return next((t[0] for t in tagTypes if name in t[1]), None)

#     @staticmethod
#     def set_columns_for_type(type, column_names):
#         Tag._column_names[type] = column_names

class Tag:
    _column_names = {}

    def __init__(self, tag_type, columns, values):
        self.tag_type = tag_type

        # Create a dictionary of { column[i]: value[i] }
        self.value_dict = dict(zip(columns, values))
        
        # Keep track of column names for each tag type used
        if tag_type not in Tag._column_names.keys():
            Tag._column_names[tag_type] = columns

    # Not yet fully implemented
    # The idea is to infer the tag type from the tag shortname
    # This will make creating tags easier, as you can simply specify the tag name and infer the type
    # The "config" for this should be stored in a file, so different naming schemes / type schemes can be loaded in
    @staticmethod
    def assumed_type_ab(name):
        tagTypes =  [    
            ("AB_AI", [ "LT", "LIT", "AIT", "FIT", "PIT", "TT", "WIT", "TT", "ZA", "ZS" ] ),
            ("AB_DA", [ "TAH", "TAL", "LAHH", "SD", "LS", "VAH", "PAH", "PAL", "FAL", "LAL", "LAH" ] ),
            ("AB_FV", [ "FV" ] ),
            ("AB_FCV", [ "FCV" ] ),
            ("AB_MOTOR", [ "P", "CP", "SC", "BL", "CF" ] ),
            ("AB_TOTALIZER", [ ] )]
        name = name.split('\\\\')[-1]
        name = name.split('_')[0]
        return next((t[0] for t in tagTypes if name in t[1]), None)



class DBConnection:
    def __init__(self, file):
        self.filename = file
        self.conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + file + ';')
        self.table_names = [row.table_name for row in self.conn.cursor().tables() if row.table_type == 'TABLE']
        self.tables = {}

    def close(self):
        self.conn.close()
    
    def get_all(self):
        cursor = self.conn.cursor()
        for table in self.table_names:
            cursor.execute(f"select * from {table}")
            self.tables[table] = [row for row in cursor]
        return self.tables
        
    def get(self, type = None):
        cursor = self.conn.cursor()
        if type == None:
            for table in self.table_names:
                cursor.execute(f"select * from {table}")
                self.tables[table] = [row for row in cursor]
            return self.tables
        else:
            cursor.execute(f"select * from {type}")
            self.tables[type] = [row for row in cursor]
            return self.tables[type]

    # Todo: This does not currently add the tags into the instance's working tables
    def add_tags(self, type, tag_list, remove_id_info = True):
        cursor = self.conn.cursor()
        cursor.execute(f"select * from {type}")
        column_names = [column[0] for column in cursor.description]
        for tag in tag_list:
            if remove_id_info:
                tag[0] = None
                tag[-1] = None
                tag[-2] = None
            query = f"insert into {type} ({','.join([f'[{x}]' for x in column_names])}) values ({','.join(['?' for x in tag])})"
            cursor.execute(query, tag)
        cursor.commit()
        return True

    def update_tags(self, type, tag_list):
        cursor = self.conn.cursor()
        cursor.execute(f"select * from {type}")
        column_names = [column[0] for column in cursor.description]
        for tag in tag_list:
            select_query = f'select * from {type} where [Export Info - leave blank for new records] = ?'
            existing_tag = cursor.execute(select_query, [tag[0]]).fetchone()
            if existing_tag == None:
                raise Exception("Update failed - tag not found.")

            updated_id = tag[0]
            updated_columns = [column_names[i] for i in range(1, len(column_names)) if tag[i] != None and tag[i] != '']
            updated_values = [tag[i] for i in range(1, len(column_names)) if tag[i] != None and tag[i] != '']
            
            query = f"update {type} set {','.join([f'[{updated_columns[i]}]=?' for i in range(len(updated_columns))])} where [Export Info - leave blank for new records] = '{updated_id}'"
            cursor.execute(query, updated_values)
        cursor.commit()
        return True

    def get_tag_by_name(self, type, name):
        cursor = self.conn.cursor()
        name = name.split('\\\\')[-1]
        cursor.execute(f"select * from {type} where Name like '%' + ?", [name])
        return cursor.fetchone()

    def get_columns_by_type(self, type):
        cursor = self.conn.cursor()
        cursor.execute(f"select * from {type}")
        return [column[0] for column in cursor.description]
