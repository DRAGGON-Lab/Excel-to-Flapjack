import os
import pandas as pd
from flapjack import Flapjack

class X2F:

    """
        inputData retrives information from input XDC excel template, connects to fj, and formats and uploads data
        includes:
            - class definition
            - initialization method
            - class attributes
            - instance attributes
            - methods


        Instance Attributes
        ----------
        fj : flapjack.flapjack.Flapjack object
            fj case that contains the connection to fj account

        excel_path: path or list of paths
            path(s) to the XDC spreadsheet(s). Distributed templates split the
            objects across several workbooks (e.g. a Study book, a Strain book,
            a SampleDesign book); a single master workbook still works.

        df : pandas dataframe
            contains all the information from XDC file in a format that can be used by FJ

        hash_map: hash map
            key: name of study attributes from XDC spreadsheet
            value: corresponding flapjackID

        fj_conv_sht: dataframe
            dataframe containing information from every workbook's flapjack_cols page
            - Sheet Name: name of the sheet (tab) that holds the data
            - FlapjackObject: the Flapjack object type the sheet maps to (new,
              distributed layout only; the legacy master reuses Sheet Name)
            - ColName: name of the column on 'Sheet Name' holding the data
            - FlapjackName: the Flapjack field name, OR a directive telling the
              uploader to resolve the column to a SynBioHub URI (see URI_DIRECTIVES)

        sbh_uri_map: dict
            name (lower-cased) -> SynBioHub URI, merged from every SBH_* collection
            sheet across all workbooks. This is the join between the template's
            names (strains, medias, chassis, plasmids, chemicals, ...) and the
            SynBioHub URIs that Flapjack uses to identify those objects.
    """
    # list of all the sheet names in a legacy master XDC excel file. Each will
    # define an fj object with data extracted from the excel sheet. In the
    # distributed layout the object types come from the FlapjackObject column
    # of flapjack_cols instead. Use self.types to access this class attribute.
    types = ['Chemical', 'DNA', 'Supplement', 'Vector', 'Strain', 'Media',
            'Signal', 'Study', 'Assay', 'Sample', 'Measurement', 'Sample Design']

    # FlapjackName values in flapjack_cols that are *directives* (an instruction
    # about what to do with a column) rather than a literal Flapjack field name.
    # They tell the uploader to resolve the column's values to SynBioHub URIs and
    # build/link Flapjack objects from them. Matched case-insensitively.
    URI_DIRECTIVES = {
        'get the sbh uri',                                 # SampleDesign: Strains/Medias/Supplements
        'add this to sample',                              # Strain: Chassis/Plasmids -> vector on the sample
        'go and grab all the sbh uris for this sample',    # Sample: pull its design's resolved URIs
    }


    def __init__(self,
                 excel_path,
                 fj_url,
                 fj_user=None,
                 fj_pass=None,
                 fj_token=None):
        # Default to the local Flapjack dev port only for bare localhost.
        if fj_url in ("localhost", "127.0.0.1"):
            fj_url = fj_url + ":8000"

        self.fj = Flapjack(url_base=fj_url)
        # Log in by token if provided, otherwise username/password.
        if fj_token:
            self.fj.log_in_token(username=fj_user, access_token=None, refresh_token=fj_token)
            self.fj.refresh()
        elif fj_user and fj_pass:
            self.fj.log_in(username=fj_user, password=fj_pass)
        else:
            raise ValueError("Flapjack credentials required: provide fj_token, or fj_user and fj_pass")

        # Load the workbook(s), merge their flapjack_cols mapping, and build the
        # name -> SynBioHub URI lookup from every SBH_* collection sheet.
        self._load_workbooks(excel_path)

        self.df = pd.DataFrame()
        self.hash_map = {}
        self.del_map = {}

        # Populated while uploading the distributed templates:
        # object type -> list of (original_column, directive) needing URI work.
        self.uri_cols = {}
        # design id (lower) -> {'strain','media','vector','supplements'} of FJ ids
        self.design_map = {}
        # strain name (lower) -> FJ vector id built from its chassis/plasmids
        self.strain_vector = {}
        # strain name (lower) -> FJ strain id
        self.strain_obj_map = {}
        # supplement id (lower) -> FJ supplement id
        self.supplement_map = {}
        # (model, name-lower, uri) -> FJ id, so a name/URI is only created once
        self._obj_cache = {}

    def _load_workbooks(self, excel_path):
        """Open every workbook and merge the flapjack_cols + SBH_* sheets.

        Accepts a single path or a list of paths. A single master workbook is
        just a one-element list, so the legacy layout keeps working.
        """
        excel_paths = list(excel_path) if isinstance(excel_path, (list, tuple)) else [excel_path]
        self.excel_paths = excel_paths
        self.xls_list = [pd.ExcelFile(p) for p in excel_paths]
        self.xls = self.xls_list[0]  # back-compat: the first/only workbook

        # Merge the column-mapping sheet from every workbook. It is named
        # 'flapjack_cols' in the distributed templates and 'FlapjackCols' in the
        # legacy master; _find_conv_sheet tolerates the case/underscore variants.
        conv_parts = []
        for xls in self.xls_list:
            conv_sheet = self._find_conv_sheet(xls)
            if conv_sheet is not None:
                conv_parts.append(xls.parse(conv_sheet, skiprows=0))
        if conv_parts:
            self.fj_conv_sht = pd.concat(conv_parts, ignore_index=True).dropna(how='all')
        else:
            self.fj_conv_sht = pd.DataFrame(columns=['Sheet Name', 'ColName', 'FlapjackName'])

        # The distributed layout is identified by the extra FlapjackObject column.
        self.distributed = 'FlapjackObject' in self.fj_conv_sht.columns

        # name -> SynBioHub URI, from every SBH_* collection sheet.
        self.sbh_uri_map = self._load_sbh_uri_map()

    @staticmethod
    def _find_conv_sheet(xls):
        """Return the flapjack_cols mapping sheet name, tolerating variants.

        Handles 'flapjack_cols' (distributed) and 'FlapjackCols' (legacy master).
        Returns None if the workbook has no mapping sheet (e.g. an SBH-only book).
        """
        for sheet in xls.sheet_names:
            if sheet.lower().replace('_', '').replace(' ', '') == 'flapjackcols':
                return sheet
        return None

    def _find_sheet_source(self, sheet_name):
        """Return the ExcelFile (of the loaded workbooks) that holds sheet_name."""
        for xls in self.xls_list:
            if sheet_name in xls.sheet_names:
                return xls
        return None

    @staticmethod
    def _find_header_row(xls, sheet_name, id_col, max_scan=6):
        """Locate a data sheet's header row by finding its '<Object> ID' column.

        Distributed templates put the header on row 0; the legacy master has a
        3-row blank preamble (header on row 3). Scanning for the ID column
        handles both without a hard-coded skiprows.
        """
        raw = xls.parse(sheet_name, header=None, nrows=max_scan)
        for i in range(len(raw)):
            values = [str(v).strip() for v in raw.iloc[i].tolist()]
            if id_col in values:
                return i
        return 0

    def _load_sbh_uri_map(self):
        """Merge every SBH_* collection sheet into a name -> URI dict."""
        uri_map = {}
        for xls in self.xls_list:
            for sheet in xls.sheet_names:
                if not sheet.lower().startswith('sbh_'):
                    continue
                df = xls.parse(sheet)
                name_col = uri_col = None
                for c in df.columns:
                    cl = str(c).strip().lower()
                    if cl == 'name':
                        name_col = c
                    elif cl in ('uri', 'sboluri', 'sbh uri', 'sbh_uri'):
                        uri_col = c
                if name_col is None or uri_col is None:
                    continue
                for _, r in df.iterrows():
                    nm, uri = r[name_col], r[uri_col]
                    if pd.notna(nm) and pd.notna(uri):
                        uri_map[str(nm).strip().lower()] = str(uri).strip()
        return uri_map

    def _resolve_uri(self, name):
        """Return the SynBioHub URI for a template name, or None if unknown."""
        if name is None or (isinstance(name, float) and pd.isna(name)):
            return None
        return self.sbh_uri_map.get(str(name).strip().lower())

    @staticmethod
    def _split_names(value):
        """Split a possibly comma-separated cell into a list of trimmed names."""
        if value is None or (isinstance(value, float) and pd.isna(value)):
            return []
        return [p.strip() for p in str(value).split(',') if p.strip()]

    @staticmethod
    def _clean(value):
        """Return None for NaN/blank so it is omitted from the create payload."""
        if value is None or (isinstance(value, float) and pd.isna(value)):
            return None
        return value

    @staticmethod
    def _nonblank(value, fallback):
        """Return value if it is a non-blank string, else fallback.

        Flapjack requires several text fields (e.g. description, machine) to be
        present and non-empty; the tutorial data leaves them blank, so we fall
        back to a sensible default (usually the object's own name).
        """
        if value is None or (isinstance(value, float) and pd.isna(value)):
            return fallback
        s = str(value).strip()
        return s if s else fallback

    def _create_or_get(self, model, name, uri=None, **extra):
        """Create a Flapjack object, tagging it with its SynBioHub URI when known.

        Results are cached by (model, name, uri) so a name/URI referenced by
        several designs is only created once. The server contract for `sboluri`
        varies by model, so if a create that includes sboluri fails we retry once
        without it — the object is still created (just untagged) and continues.
        """
        cache_key = (model, str(name).strip().lower(), uri or '')
        if cache_key in self._obj_cache:
            return self._obj_cache[cache_key]

        base = dict(name=name)
        base.update({k: v for k, v in extra.items() if v is not None})
        attempts = []
        if uri:
            attempts.append(dict(base, sboluri=uri))
        attempts.append(base)
        for kwargs in attempts:
            try:
                fj_obj = self.fj.create(model, confirm=False, overwrite=False, **kwargs)
                fid = fj_obj.id[0]
            except Exception as e:
                print(f"    [create {model}] '{name}' failed with {list(kwargs)}: {type(e).__name__}: {e}")
                continue
            self.del_map.setdefault(model, [])
            if fid not in self.del_map[model]:
                self.del_map[model].append(fid)
            self._obj_cache[cache_key] = fid
            return fid
        return None

    def _build_vector(self, name, part_names):
        """Build one Flapjack vector from a strain's DNA parts (chassis/plasmids).

        DNA and vector are treated as a single logical object: each part becomes
        a dna tagged with its SynBioHub URI, and they are bundled into one vector
        that gets attached to the sample.
        """
        dna_ids = []
        for part in part_names:
            did = self._create_or_get('dna', part, uri=self._resolve_uri(part))
            if did is not None:
                dna_ids.append(did)
        if not dna_ids:
            return None
        try:
            vec = self.fj.create('vector', confirm=False, overwrite=False,
                                  name=name, dnas=dna_ids)
            vid = vec.id[0]
        except Exception as e:
            print(f"    [create vector] '{name}' failed: {type(e).__name__}: {e}")
            return None
        self.del_map.setdefault('vector', [])
        if vid not in self.del_map['vector']:
            self.del_map['vector'].append(vid)
        return vid

    # def fj_login(fj_user, fj_pass):
        # self.fj.log_in(username = fj_user, password=fj_pass)


    def print_info(self, fj_sht = False, df=False, hash_map=False):
        if fj_sht:
            print(self.fj_conv_sht)
        if df:
            print(self.df.head())
            print(self.df.info())
        if hash_map:
            print(self.hash_map)

    def create_df(self, header_skiprows=None, skip_objects=('Measurement',)):
        """Parse every mapped sheet across all workbooks into self.df.

        Uses the flapjack_cols mapping to decide, per (Sheet Name, FlapjackObject)
        pair, which columns to keep and how to rename them. Columns whose
        FlapjackName is a URI directive are kept under their original header (not
        renamed) and recorded in self.uri_cols for the SynBioHub-URI resolution
        step. header_skiprows=None auto-detects the header row.
        """
        conv = self.fj_conv_sht
        if conv.empty:
            return

        has_obj = 'FlapjackObject' in conv.columns
        if has_obj:
            pairs = conv[['Sheet Name', 'FlapjackObject']].dropna().drop_duplicates().values.tolist()
        else:
            # legacy master: the object type IS the sheet name
            pairs = [(s, s) for s in conv['Sheet Name'].dropna().unique()]

        for sheet_name, fj_obj in pairs:
            if fj_obj in skip_objects:
                continue
            xls = self._find_sheet_source(sheet_name)
            if xls is None:
                print(f"[create_df] sheet '{sheet_name}' not found in any workbook; skipping {fj_obj}")
                continue

            id_col = f"{str(sheet_name).title()} ID"
            hdr = self._find_header_row(xls, sheet_name, id_col) if header_skiprows is None else header_skiprows
            try:
                obj_df = xls.parse(sheet_name, skiprows=hdr, index_col=id_col)
            except (ValueError, KeyError) as e:
                print(f"[create_df] could not index '{sheet_name}' by '{id_col}': {e}; skipping {fj_obj}")
                continue

            grp = conv[conv['Sheet Name'] == sheet_name]
            if has_obj:
                grp = grp[grp['FlapjackObject'] == fj_obj]

            rename = {}
            uri_cols = []
            seen_targets = set()
            for _, r in grp.iterrows():
                col, fj_name = r['ColName'], r['FlapjackName']
                if pd.isna(col) or pd.isna(fj_name):
                    continue
                if col not in obj_df.columns:
                    print(f"[create_df] {fj_obj}: flapjack_cols maps '{col}' but the "
                          f"'{sheet_name}' tab has no such column {list(obj_df.columns)}; "
                          f"mapping skipped")
                    continue
                if str(fj_name).strip().lower() in self.URI_DIRECTIVES:
                    if (col, str(fj_name).strip().lower()) not in uri_cols:
                        uri_cols.append((col, str(fj_name).strip().lower()))
                elif fj_name not in seen_targets:
                    rename[col] = fj_name
                    seen_targets.add(fj_name)

            keep = list(rename.keys()) + [c for c, _ in uri_cols]
            obj_df = obj_df[keep].rename(columns=rename)
            obj_df['object'] = fj_obj
            obj_df['flapjackid'] = ''

            if uri_cols:
                self.uri_cols.setdefault(fj_obj, [])
                for uc in uri_cols:
                    if uc not in self.uri_cols[fj_obj]:
                        self.uri_cols[fj_obj].append(uc)

            self.df = pd.concat([self.df, obj_df])


    def upload_studies(self, overwrite=False, confirm=False):
        del_inds = []
        for index, row in self.df[self.df['object'] == 'Study'].iterrows():
            fj_obj = self.fj.create(
                'study',
                name=row['name'],
                description=self._nonblank(row.get('description'), row['name']),
                confirm=confirm,
                overwrite=overwrite,
            )
            for i in fj_obj:
                print(i, fj_obj[i][0])
            print()
            self.hash_map[index] = fj_obj.id[0]
            if fj_obj.id[0] not in del_inds:
                del_inds.append(fj_obj.id[0])
            self.df.loc[index, 'flapjackid'] = fj_obj.id[0]
            # if row[DOI] is not nan patch the DOI
            if 'DOI' in row and not pd.isna(row['DOI']):
                self.fj.patch('study', fj_obj.id[0], doi=row['DOI'])
        self.del_map['study'] = del_inds


    def upload_signals(self, overwrite=False, confirm=False):
        del_inds = []
        for index, row in self.df[self.df['object'] == 'Signal'].iterrows():
            fj_obj = self.fj.create(
                'signal',
                name=row['name'],
                description=row['description'],
                color=row['color'],
                confirm=confirm,
                overwrite=overwrite,
            )
            for i in fj_obj:
                print(i, fj_obj[i][0])
            print()
            self.hash_map[index] = fj_obj.id[0]
            if fj_obj.id[0] not in del_inds:
                del_inds.append(fj_obj.id[0])
            self.df.loc[index, 'flapjackid'] = fj_obj.id[0]
        self.del_map['signal'] = del_inds


    def upload_chemicals(self, overwrite=False, confirm=False):
        del_inds = []
        for index, row in self.df[self.df['object'] == 'Chemical'].iterrows():
            fj_obj = self.fj.create(
                'chemical',
                name=row['name'],
                description=row['description'],
                confirm=confirm,
                overwrite=overwrite,
            )
            for i in fj_obj:
                print(i, fj_obj[i][0])
            print()
            self.hash_map[index] = fj_obj.id[0]
            if fj_obj.id[0] not in del_inds:
                del_inds.append(fj_obj.id[0])
            self.df.loc[index, 'flapjackid'] = fj_obj.id[0]
        self.del_map['chemical'] = del_inds

    def upload_dna(self, overwrite=False, confirm=False):
        del_inds = []
        for index, row in self.df[self.df['object'] == 'DNA'].iterrows():
            fj_obj = self.fj.create(
                'dna',
                name=row['name'],
                confirm=confirm,
                overwrite=overwrite,
            )
            for i in fj_obj:
                print(i, fj_obj[i][0])
            print()
            self.hash_map[index] = fj_obj.id[0]
            if fj_obj.id[0] not in del_inds:
                del_inds.append(fj_obj.id[0])
            self.df.loc[index, 'flapjackid'] = fj_obj.id[0]
        self.del_map['dna'] = del_inds


    def upload_medias(self, overwrite=False, confirm=False):
        del_inds = []
        for index, row in self.df[self.df['object'] == 'Media'].iterrows():
            fj_obj = self.fj.create(
                'media',
                name=row['name'],
                description=row['description'],
                confirm=confirm,
                overwrite=overwrite,
            )
            for i in fj_obj:
                print(i, fj_obj[i][0])
            print()
            self.hash_map[index] = fj_obj.id[0]
            if fj_obj.id[0] not in del_inds:
                del_inds.append(fj_obj.id[0])
            self.df.loc[index, 'flapjackid'] = fj_obj.id[0]
        self.del_map['media'] = del_inds


    def upload_strains(self, overwrite=False, confirm=False):
        del_inds = []
        for index, row in self.df[self.df['object'] == 'Strain'].iterrows():
            fj_obj = self.fj.create(
                'strain',
                name=row['name'],
                description=row['description'],
                confirm=confirm,
                overwrite=overwrite,
            )
            for i in fj_obj:
                print(i, fj_obj[i][0])
            print()
            self.hash_map[index] = fj_obj.id[0]
            if fj_obj.id[0] not in del_inds:
                del_inds.append(fj_obj.id[0])
            self.df.loc[index, 'flapjackid'] = fj_obj.id[0]
        self.del_map['strain'] = del_inds


    def upload_supplements(self, overwrite=False, confirm=False):
        del_inds = []
        for index, row in self.df[self.df['object'] == 'Supplement'].iterrows():
            fj_obj = self.fj.create(
                'supplement',
                name=row['name'],
                description=row['description'],
                chemical=self.df.loc[row['chemical'], 'flapjackid'],
                concentration=row['concentration'],
                confirm=confirm,
                overwrite=overwrite,
            )
            for i in fj_obj:
                print(i, fj_obj[i][0])
            print()
            self.hash_map[index] = fj_obj.id[0]
            if fj_obj.id[0] not in del_inds:
                del_inds.append(fj_obj.id[0])
            self.df.loc[index, 'flapjackid'] = fj_obj.id[0]
        self.del_map['supplement'] = del_inds


    def upload_vectors(self, overwrite=False, confirm=False):
        del_inds = []
        for index, row in self.df[self.df['object'] == 'Vector'].iterrows():
            # print(self.df.loc[row['dna'], 'flapjackid'])
            fj_obj = self.fj.create(
                'vector',
                name=row['name'],
                description=(row['description'] if pd.notna(row['description']) else ''),
                dnas=[self.df.loc[row['dnas'], 'flapjackid']],
                confirm=confirm,
                overwrite=overwrite,
            )
            for i in fj_obj:
                print(i, fj_obj[i][0])
            print()
            self.hash_map[index] = fj_obj.id[0]
            if fj_obj.id[0] not in del_inds:
                del_inds.append(fj_obj.id[0])
            self.df.loc[index, 'flapjackid'] = fj_obj.id[0]
        self.del_map['vector'] = del_inds


    def upload_assays(self, overwrite=False, confirm=False):
        del_inds = []
        for index, row in self.df[self.df['object'] == 'Assay'].iterrows():
            fj_obj = self.fj.create(
                'assay',
                study=self.df.loc[row['study'], 'flapjackid'],
                name=row['name'],
                description=self._nonblank(row.get('description'), row['name']),
                machine=self._nonblank(row.get('machine'), 'unspecified'),
                temperature=self._clean(row.get('temperature')),
                confirm=confirm,
                overwrite=overwrite,
            )
            for i in fj_obj:
                print(i, fj_obj[i][0])
            print()
            self.hash_map[index] = fj_obj.id[0]
            if fj_obj.id[0] not in del_inds:
                del_inds.append(fj_obj.id[0])
            self.df.loc[index, 'flapjackid'] = fj_obj.id[0]
        self.del_map['assay'] = del_inds


    def upload_samples(self, overwrite=False, confirm=False):
        del_inds = []
        # get fj ids from sample design
        for index, row in self.df[self.df['object'] == 'Sample'].iterrows():
            media_index = self.df.loc[row['sampledesign'], 'media']
            media_id = self.df.loc[media_index, 'flapjackid']
            strain_index = self.df.loc[row['sampledesign'], 'strain']
            strain_id = self.df.loc[strain_index, 'flapjackid']
            vector_index = self.df.loc[row['sampledesign'], 'vector']
            vector_id = self.df.loc[vector_index, 'flapjackid']

            fj_obj = self.fj.create(
                'sample',
                assay=self.df.loc[row['assay'], 'flapjackid'],
                media=media_id,
                strain=strain_id,
                vector=vector_id,
                row=row['row'],
                col=row['col'],
                confirm=confirm,
                overwrite=overwrite,
            )
            for i in fj_obj:
                print(i, fj_obj[i][0])
            print()
            self.hash_map[index] = fj_obj.id[0]
            if fj_obj.id[0] not in del_inds:
                del_inds.append(fj_obj.id[0])
            self.df.loc[index, 'flapjackid'] = fj_obj.id[0]

            if not pd.isna(row['supplement']):
                supplement_index = self.df.loc[row['sampledesign'], 'supplement']
                supplement_id = self.df.loc[supplement_index, 'flapjackid']
                self.fj.patch('sample', fj_obj.id[0], supplements=supplement_id)
        self.del_map['sample'] = del_inds


    # ------------------------------------------------------------------
    # Distributed-template upload (SynBioHub-URI driven)
    # ------------------------------------------------------------------

    def upload_all(self):
        if self.df.empty:
            self.create_df()
        if self.distributed:
            self._upload_distributed()
        else:
            self._upload_legacy()

    def _run_step(self, label, fn):
        """Run one upload step, reporting failures without aborting the run."""
        try:
            fn()
            print(f"[upload] {label}: OK")
        except Exception as e:
            import traceback
            print(f"[upload] {label}: FAILED ({type(e).__name__}: {e})")
            traceback.print_exc()

    def _present(self, obj):
        return 'object' in self.df.columns and (self.df['object'] == obj).any()

    def _upload_legacy(self):
        """Original single-master flow, now resilient to per-type failures."""
        steps = [
            ('Study', self.upload_studies), ('Signal', self.upload_signals),
            ('Chemical', self.upload_chemicals), ('DNA', self.upload_dna),
            ('Media', self.upload_medias), ('Strain', self.upload_strains),
            ('Supplement', self.upload_supplements), ('Vector', self.upload_vectors),
            ('Assay', self.upload_assays), ('Sample', self.upload_samples),
        ]
        for obj, fn in steps:
            if self._present(obj):
                self._run_step(obj, fn)

    def _upload_distributed(self):
        """SynBioHub-URI driven upload for the distributed templates.

        Order matters: studies before assays (assay -> study), supplements and
        strain-vectors before designs, designs before samples.
        """
        if self._present('Study'):
            self._run_step('Study', self.upload_studies)
        if self._present('Signal'):
            self._run_step('Signal', self.upload_signals)
        if self._present('Assay'):
            self._run_step('Assay', self.upload_assays)
        if self._present('Supplement'):
            self._run_step('Supplement', self._up_supplements_dist)
        if self._present('Strain'):
            self._run_step('Strain vectors', self._up_strain_vectors_dist)
        if self._present('SampleDesign'):
            self._run_step('SampleDesign', self._up_designs_dist)
        if self._present('Sample'):
            self._run_step('Sample', self._up_samples_dist)

    def _up_supplements_dist(self):
        """Create each supplement's chemical (SBH-URI tagged) then the supplement."""
        for index, row in self.df[self.df['object'] == 'Supplement'].iterrows():
            chem_id = None
            chem_name = row.get('chemical')
            if pd.notna(chem_name):
                chem_id = self._create_or_get('chemical', str(chem_name),
                                              uri=self._resolve_uri(chem_name),
                                              description=str(chem_name))
            try:
                fj_obj = self.fj.create('supplement', confirm=False, overwrite=False,
                                        name=row['name'], chemical=chem_id,
                                        description=self._nonblank(row.get('description'), row['name']),
                                        concentration=self._clean(row.get('concentration')))
                sid = fj_obj.id[0]
            except Exception as e:
                print(f"    [create supplement] '{row.get('name')}' failed: {type(e).__name__}: {e}")
                continue
            self.supplement_map[str(index).lower()] = sid
            self.df.loc[index, 'flapjackid'] = sid
            self.del_map.setdefault('supplement', [])
            if sid not in self.del_map['supplement']:
                self.del_map['supplement'].append(sid)

    def _up_strain_vectors_dist(self):
        """For each strain, create the strain object and bundle its DNA parts.

        The strain sheet's Chassis/Plasmid columns are flagged 'add this to
        sample'; they become one vector (via _build_vector) that the sample will
        reference. The strain itself is tagged with its SynBioHub URI.
        """
        for index, row in self.df[self.df['object'] == 'Strain'].iterrows():
            strain_name = row['name'] if ('name' in row and pd.notna(row.get('name'))) else str(index)
            uri = self._resolve_uri(strain_name) or self._resolve_uri(index)
            desc = self._nonblank(row.get('description'), str(strain_name))
            sid = self._create_or_get('strain', str(strain_name), uri=uri, description=desc)
            if sid is not None:
                self.strain_obj_map[str(strain_name).lower()] = sid
                self.df.loc[index, 'flapjackid'] = sid

            parts = []
            for col, _ in self.uri_cols.get('Strain', []):
                parts.extend(self._split_names(row.get(col)))
            vid = self._build_vector(str(strain_name), parts)
            if vid is not None:
                self.strain_vector[str(strain_name).lower()] = vid

    def _up_designs_dist(self):
        """Resolve each sample design's components to Flapjack ids via SBH URIs."""
        for index, row in self.df[self.df['object'] == 'SampleDesign'].iterrows():
            entry = {'strain': None, 'media': None, 'vector': None, 'supplements': []}
            for col, _ in self.uri_cols.get('SampleDesign', []):
                names = self._split_names(row.get(col))
                if not names:
                    continue
                cl = col.lower()
                if 'strain' in cl:
                    nm = names[0]
                    entry['strain'] = (self.strain_obj_map.get(nm.lower())
                                       or self._create_or_get('strain', nm, uri=self._resolve_uri(nm), description=nm))
                    entry['vector'] = self.strain_vector.get(nm.lower())
                elif 'media' in cl:
                    nm = names[0]
                    entry['media'] = self._create_or_get('media', nm, uri=self._resolve_uri(nm), description=nm)
                elif 'supplement' in cl:
                    for nm in names:
                        supp = self.supplement_map.get(nm.lower())
                        if supp is not None:
                            entry['supplements'].append(supp)
            self.design_map[str(index).lower()] = entry

    def _up_samples_dist(self):
        """Create each well as a Flapjack sample, linking its design's objects.

        The sample's design reference is matched against design_map. Wells whose
        design cannot be resolved (a known template/data gap) are skipped and
        reported rather than aborting the run.
        """
        skipped = 0
        for index, row in self.df[self.df['object'] == 'Sample'].iterrows():
            design_ref = row.get('sampledesign')
            entry = self.design_map.get(str(design_ref).strip().lower()) if pd.notna(design_ref) else None
            if entry is None:
                skipped += 1
                print(f"    [sample] well '{index}': no design match for '{design_ref}'; skipping")
                continue

            assay_ref = row.get('assay')
            assay_id = None
            if pd.notna(assay_ref) and assay_ref in self.df.index:
                assay_id = self.df.loc[assay_ref, 'flapjackid']

            # Flapjack treats strain/vector as lists (co-culture support) and
            # requires at least one strain; control wells with no strain (e.g. a
            # media blank) cannot be created.
            if entry['strain'] is None:
                skipped += 1
                print(f"    [sample] well '{index}': design '{design_ref}' has no strain "
                      f"(control well); Flapjack requires >=1 strain, skipping")
                continue
            kwargs = dict(assay=assay_id, media=entry['media'], strain=[entry['strain']],
                          row=self._clean(row.get('row')), col=self._clean(row.get('col')))
            if entry['vector'] is not None:
                kwargs['vector'] = [entry['vector']]
            try:
                fj_obj = self.fj.create('sample', confirm=False, overwrite=False, **kwargs)
                sid = fj_obj.id[0]
            except Exception as e:
                print(f"    [sample] well '{index}' failed: {type(e).__name__}: {e}")
                continue
            self.df.loc[index, 'flapjackid'] = sid
            self.del_map.setdefault('sample', [])
            if sid not in self.del_map['sample']:
                self.del_map['sample'].append(sid)
            if entry['supplements']:
                self.fj.patch('sample', sid, supplements=entry['supplements'])
        if skipped:
            print(f"    [sample] {skipped} well(s) skipped (unresolved design reference)")

    def delete_all(self):
        for model in self.del_map:
            for id in self.del_map[model]:
                self.fj.delete(model, id, confirm=False)
