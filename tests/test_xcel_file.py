import unittest
from unittest.mock import MagicMock, patch
from excel2flapjack.main import X2F
import pandas as pd


data = [
    {"Sheet Name": "Supplement", "ColName": "Supplement Name", "FlapjackName": "name"},
    {"Sheet Name": "Supplement", "ColName": "Chemical ID", "FlapjackName": "chemical"},
    {"Sheet Name": "Supplement", "ColName": "Concentration", "FlapjackName": "concentration"},
    {"Sheet Name": "Chemical", "ColName": "Chemical Name", "FlapjackName": "name"},
    {"Sheet Name": "Chemical", "ColName": "Chemical Description", "FlapjackName": "description"},
    {"Sheet Name": "Chemical", "ColName": "Pubchem ID", "FlapjackName": "pubchemid"},
    {"Sheet Name": "Vector", "ColName": "Vector Name", "FlapjackName": "name"},
    {"Sheet Name": "Vector", "ColName": "DNA ID", "FlapjackName": "dna"},
    {"Sheet Name": "Strain", "ColName": "Strain Name", "FlapjackName": "name"},
    {"Sheet Name": "Strain", "ColName": "Strain Description", "FlapjackName": "description"},
    {"Sheet Name": "Strain", "ColName": "Strain ID", "FlapjackName": "strain"},
    {"Sheet Name": "Media", "ColName": "Media Name", "FlapjackName": "name"},
    {"Sheet Name": "Media", "ColName": "Media Description", "FlapjackName": "description"},
    {"Sheet Name": "Signal", "ColName": "Signal Name", "FlapjackName": "name"},
    {"Sheet Name": "Signal", "ColName": "Signal Description", "FlapjackName": "description"},
    {"Sheet Name": "Signal", "ColName": "Signal Color", "FlapjackName": "color"},
    {"Sheet Name": "DNA", "ColName": "DNA Name", "FlapjackName": "name"},
    {"Sheet Name": "Measurement", "ColName": "Signal ID", "FlapjackName": "signal"},
    {"Sheet Name": "Measurement", "ColName": "Value", "FlapjackName": "value"},
    {"Sheet Name": "Measurement", "ColName": "Sample ID", "FlapjackName": "sample"},
    {"Sheet Name": "Measurement", "ColName": "Time", "FlapjackName": "time"},
    {"Sheet Name": "Sample", "ColName": "Row", "FlapjackName": "row"},
    {"Sheet Name": "Sample", "ColName": "Column", "FlapjackName": "col"},
    {"Sheet Name": "Sample", "ColName": "Sample Design ID", "FlapjackName": "sampledesign"},
    {"Sheet Name": "Sample", "ColName": "Assay ID", "FlapjackName": "assay"},
    {"Sheet Name": "Sample Design", "ColName": "Supplement ID", "FlapjackName": "supplement"},
    {"Sheet Name": "Sample Design", "ColName": "Vector ID", "FlapjackName": "vector"},
    {"Sheet Name": "Sample Design", "ColName": "Strain ID", "FlapjackName": "strain"},
    {"Sheet Name": "Sample Design", "ColName": "Media ID", "FlapjackName": "media"},
    {"Sheet Name": "Assay", "ColName": "Assay Name", "FlapjackName": "name"},
    {"Sheet Name": "Assay", "ColName": "Machine", "FlapjackName": "machine"},
    {"Sheet Name": "Assay", "ColName": "Description", "FlapjackName": "description"},
    {"Sheet Name": "Assay", "ColName": "Study ID", "FlapjackName": "study"},
    {"Sheet Name": "Assay", "ColName": "Temperature", "FlapjackName": "temperature"},
    {"Sheet Name": "Study", "ColName": "Study Name", "FlapjackName": "name"},
    {"Sheet Name": "Study", "ColName": "Description", "FlapjackName": "description"},
    {"Sheet Name": "Study", "ColName": "DOI", "FlapjackName": "DOI"},
    {"Sheet Name": "Study", "ColName": None, "FlapjackName": "public"},
]

# Create the DataFrame
df = pd.DataFrame(data)

print(df)
