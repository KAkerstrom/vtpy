import pyodbc

class Tag:
    """A class to represent a VTScada tag."""

    id_col = r'Export Info - leave blank for new records'
    _column_names = {}


    def __init__(self, tag_type : str, columns : list, values : list):
        """Tag constructor.
        
        Parameters
        ----------
        tag_type : str
            The name of the table this tag is found under in the tag export database.
        columns : list
            The list of columns for the table this tag is found under in the tag export database, in order.
        values : list
            The list of values for each column, in order.
        """
        self.tag_type = tag_type

        # Create a dictionary of { column[i]: value[i] }
        self.value_dict = dict(zip(columns, values))
        
        # Keep track of column names for each tag type used
        if tag_type not in Tag._column_names.keys():
            Tag._column_names[tag_type] = columns

    def set(self, column : str, value : str):
        """Set the column value for the tag.
        
        Parameters
        ----------
        column : str
            The column name to set the value for.
            You can use Tag.id_col for the Export Info column, rather than typing it all out.
        value : str
            The value being set.
        """

        self.value_dict[column] = value

    def get(self, column : str) -> str:
        """Gets the column value for the tag.
        
        Parameters
        ----------
        column : str
            The column name to get the value for.
            You can use Tag.id_col for the Export Info column, rather than typing it all out.

        Returns
        ----------
        str
            The value for the given column.
        """

        return self.value_dict[column]

    def values_as_list(self) -> list:
        """Gets the values of the tag as a list, in the order of the columns in the database.
        Mainly intended for database operations.

        Returns
        ----------
        list
            A list of the tag's values, in database order."""

        columns = Tag._column_names[self.tag_type]
        values = [self.value_dict[col] for col in columns]
        return values

    def remove_id_info(self):
        """Sets the 'Export Info', 'AuditName', and 'Original Shortname' values to empty strings.
        Useful for copying tags, or importing tags to a different database."""

        id_properties = [Tag.id_col, "AuditName", "Original Shortname"]
        for prop in id_properties:
            if prop in self.value_dict.keys():
                self.set(prop, '')

    def columns(self) -> list:
        """Gets the columns for this tag's tag type, in database order.

        Returns
        ----------
        list
            The columns for this tag type, in database order."""
        return Tag._column_names[self.tag_type]
        

    @staticmethod
    def assumed_type_ab(name : str) -> str:
        """NOT YET IMPLEMENTED
        The idea is to infer the tag type from the tag shortname
        This will make creating tags easier, as you can simply specify the tag name and infer the type
        The "config" for this should be stored in a file, so different naming schemes / type schemes can be loaded in"""
        return

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

    @staticmethod
    def separate_tags_by_type(tag_list : list) -> dict:
        """Separates a list of tags into a dictionary of { tag_type: [Tag] }
        
        Parameters
        ----------
        tag_list : str
            The list of tags to separate.

        Returns
        ----------
        dict
            A dictionary of { tag_type: [Tag] }
        """
        tags_by_type = {}
        for tag in tag_list:
            if tag.tag_type in tags_by_type.keys():
                tags_by_type[tag.tag_type].append(tag)
            else:
                tags_by_type[tag.tag_type] = [tag]
        return tags_by_type



class DBConnection:
    """A class that encapsulates a connection to a VTScada tag export database."""

    def __init__(self, file : str):
        """DBConnection constructor.
        Connects to the database automatically when instantiated.
        
        Parameters
        ----------
        file : str
            The filepath to the MS Access .mdb file.
        """
        self.filename = file
        self.conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + file + ';')

        # Create a dictionary of { table_name: columns_as_tuple }
        table_names =  (row.table_name for row in self.conn.cursor().tables() if row.table_type == 'TABLE')
        self.table_columns = dict(((table, self.get_columns_by_type(table)) for table in table_names))

    def close(self):
        """Closes the database connection.
        The connection is also closed automatically when out of scope."""
        self.conn.close()
    
    def get_tags(self, tag_type : str = None) -> list:
        """Creates and returns a list of Tag objects from the database.
        
        Parameters
        ----------
        tag_type : str, optional
            The name of the table this tag is found under in the tag export database.
            Leave as default to get all tags from every table.

        Returns
        ----------
        list
            A list of Tag objects.
        """
        if tag_type == None: # Get all tags
            cursor = self.conn.cursor()
            tags = []
            for table in self.table_columns.keys():
                cursor.execute(f"select * from {table}")
                table_tags = [Tag(table, self.table_columns[table], list(row)) for row in cursor]
                tags.extend(table_tags)
            return tags
        else: # Get tags for a specific type
            cursor = self.conn.cursor()
            cursor.execute(f"select * from {tag_type}")
            return [Tag(table, self.table_columns[table], row) for row in cursor]

    def add_tags(self, tag_list : list, remove_id_info : bool = True):
        """Appends the tags in tag_list to the database.
        
        Parameters
        ----------
        tag_list : list
            A list of Tag objects.
        remove_id_info : bool
            If True, sets the 'Export Info', 'AuditName', and 'Original Shortname' values to empty strings.
            Normally required when adding new tags to a database. (default is True)
            

        Returns
        ----------
        dict
            A dictionary of { tag_type: [Tag] }
        """
        tags_by_type = Tag.separate_tags_by_type(tag_list)

        cursor = self.conn.cursor()
        for tag_type in tags_by_type.keys():
            for tag in tags_by_type[tag_type]:
                if remove_id_info:
                    tag.remove_id_info()
                values = tag.values_as_list()
                query = f"insert into {tag_type} ({','.join([f'[{x}]' for x in Tag._column_names[tag_type]])}) values ({','.join(['?']*len(values))})"
                cursor.execute(query, values)
        cursor.commit()
        return True

    # TODO - return the # of rows updated
    def update_tags(self, tag_list : list):
        """Updates each tag in tag_list in the database, going by the Export Info field as the tag's ID.
        
        Parameters
        ----------
        tag_list : list
            A list of Tag objects.
        """
        tags_by_type = Tag.separate_tags_by_type(tag_list)

        cursor = self.conn.cursor()
        for tag_type in tags_by_type.keys():
            updated_columns = Tag._column_names[tag_type]
            for tag in tags_by_type[tag_type]:
                select_query = f'select * from {tag_type} where [Export Info - leave blank for new records] = ?'
                updated_id = tag.get(Tag.id_col)
                existing_tag = cursor.execute(select_query, [updated_id]).fetchone()
                if existing_tag == None:
                    raise Exception("Update failed - tag not found.")
                params = tag.values_as_list()
                params.append(updated_id)
                query = f"update {tag_type} set {','.join([f'[{updated_columns[i]}]=?' for i in range(len(updated_columns))])} where [Export Info - leave blank for new records] = ?"
                cursor.execute(query, params)
        cursor.commit()
        return True

    # TODO - Remove need for tag_type
    def get_tag_by_name(self, tag_type : str, name : str) -> Tag:
        """Finds a single tag in the database by its Name (or ShortName).
        
        Parameters
        ----------
        tag_type : list
            The name of the table this tag is found under in the tag export database.
        name : str
            The Name of the tag. Everything before the slashes is dropped, so this can be the full path or the Short Name.
            

        Returns
        ----------
        Tag
            The found Tag, if found. Otherwise, returns None.
        """
        cursor = self.conn.cursor()
        name = name.split('\\\\')[-1]
        cursor.execute(f"select * from {tag_type} where Name like '%' + ?", [name])
        tag = cursor.fetchone()
        return Tag(tag_type, self.table_columns[tag_type], tag) if tag != None else None

    def get_columns_by_type(self, tag_type : str) -> list:
        """Gets the columns by type.
        This method is run and cached on instantiation as the self.table_columns dict, so shouldn't need to be run manually.
        
        Parameters
        ----------
        tag_type : list
            The name of the table this tag is found under in the tag export database.
            
        Returns
        ----------
        list
            The columns for the given tag type, in database order.
        """
        cursor = self.conn.cursor()
        cursor.execute(f"select * from {tag_type}")
        if len(cursor.description) == 0:
            print (tag_type)
        return [column[0] for column in cursor.description]

    def create_tag_template(self, tag_type : str) -> Tag:
        """Creates an empty Tag object of the given tag type.
        Useful for creating and adding new tags to the database.
        
        Parameters
        ----------
        tag_type : list
            The name of the table this tag is found under in the tag export database.
            
        Returns
        ----------
        Tag
            An empty Tag object, with the tag type and columns preconfigured.
        """
        columns = self.table_columns[tag_type]
        return Tag(tag_type, columns, ['']*len(columns))

