# -*- coding: utf-8 -*-
"""
Created on Sun Feb 26 18:26:54 2023

@author: Yue
"""


decision_matrix = {
    "high_gpa": {1: "Discuss", 2: "Look Again", 3: "Wait - High", 4: "Admit", 5: "Admit"},
    "2nd_rev_if_needed": {1: "Deny", 2: "Look Again", 3: "Wait - Low", 4: "Send to 2nd Rev", 5: "Send to 2nd Rev"},
    "missing_2nd_rev": {1: "Need 2nd Rev", 2: "Need 2nd Rev", 3: "Need 2nd Rev", 4: "Need 2nd Rev", 5: "Need 2nd Rev"},
    1: {1: "Deny", 2: "Look Again", 3: "Discuss", 4: "Discuss", 5: "Discuss"},
    2: {1: "Look Again", 2: "Look Again", 3: "Look Again", 4: "Wait - High", 5: "Discuss"},
    3: {1: "Discuss", 2: "Look Again", 3: "Wait - Low", 4: "Discuss", 5: "Discuss"},
    4: {1: "Discuss", 2: "Look Again", 3: "Wait - High", 4: "Admit", 5: "Admit"},
    5: {1: "Discuss", 2: "Look Again", 3: "Discuss", 4: "Admit", 5: "Admit"},
}

