'''

The MIT License (MIT)

Copyright (c) 2021 Nestfield Co., Ltd. 
<https://www.nestfield.co.kr>             
Author: Wonseok Song

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.


The openpyxl is under the MIT/Expat license
(see https://github.com/chronossc/openpyxl/blob/master/LICENCE).

'''

import sys
import openpyxl as xl
import json

PARSER_VERSION_STRING = '2021.06.22.build-1'


# chohpower
#if len(sys.argv) != 2:
#    print("usage: xls2aas [xlsx file name]")
if len(sys.argv) != 4:
    print("usage: xls2aas [xlsx file name] [output name] [result file name]")
    sys.exit()




#-------------------- common sub functions --------------------------
VTYPE_STRING            = 'string'
VTYPE_LANG_STRING       = 'langString'

fp_result = open(sys.argv[3], mode = 'wt')


def outMessage(msg):
    print(msg)
    fp_result.write(msg)
    fp_result.write("\n")
    return


def splitMultiLangText(str):
    splitTexts = str.split('@')
    #print(splitTexts)

    dictStr = []

    for text in splitTexts:
        if len(text) < 3:
            continue

        else:
            langStr = text.split(':')
            if len(langStr) != 2:
                continue;

            item = {}
            item['language'] = langStr[0].strip().lower()
            item['text'] = langStr[1].strip()

            dictStr.append(item) 

    return dictStr

def splitOptionsText(str):
    options = []
    
    if len(str) == 0:
        return options

    items = str.split(',')
    if len(items) > 0:
        for optionText in items:
            optionText = optionText.strip()
            if len(optionText) == 0:
                continue

            optionText = optionText.split('=')
            if len(optionText) != 2:
                continue

            optionText[0] = optionText[0].strip()
            optionText[1] = optionText[1].strip()

            optionItem = {}
            optionItem[optionText[0].lower()] = optionText[1]
            options.append(optionItem)
        

    return options


def getValueObject(type, str):

    if type.lower() == VTYPE_STRING:
        return str

    elif type.lower() == VTYPE_LANG_STRING:
        return splitMultiLangText(str)

    return str



def getValueTypeText(str):

    if str.lower() == 'string':
        return VTYPE_STRING

    elif str.lower() == 'langstring':
        return VTYPE_LANG_STRING


    return str

def getkeyValueFromString(str):
    # parsing 'string [{type/local/value/index/idType} ... {}]'
    if str == None:
        return None
		
    if len(str) < 1:
        return None
	
    keys_array = []
	
    keyValues = str.split('}')
    for kv in keyValues:
        kv = kv.strip()
        kv = kv.lstrip('{')
        kv = kv.strip()
        if len(kv) < 1:
            continue
			
        kv_items = kv.split(',')
        if kv_items == None:
            continue
		
        if len(kv_items) != 5:
            continue
			
        keyData = {}
        keyData['index'] = kv_items[0]
        keyData['type'] = kv_items[1]
        keyData['idType'] = kv_items[3]
        keyData['value'] = kv_items[4]
		
        if kv_items[2].lower() == 'true':
            keyData['local'] = 'True'
			
        elif kv_items[2].lower() == 'false':
            keyData['local'] = 'False'
			
        else:
            continue;

        keys_array.append(keyData)
	
    if len(keys_array) < 1:
        return None
		
    return keys_array

#-------------------- making aas.json codes -------------------------
aasx = {}
aasx['assetAdministrationShells'] = []
aasx['assets'] = []
aasx['submodels'] = []
aasx['conceptDescriptions'] = []




def add_Shell(state, idShort, idType, idValue, refType, refLocal):
 
    dictShell = {}

    # fill default array or dictionaries (top level)
    dictShell['hasDataSpecification'] = []
    dictShell['asset'] = {}
    dictShell['submodels'] = []
    dictShell['conceptDictionaries'] = []
    dictShell['identification'] = {}
    dictShell['administration'] = {}
    dictShell['idShort'] = idShort
    dictShell['modelType'] = {}
    dictShell['descriptions'] = []

    # fill default array or dictionaries (2nd level)
    dictShell['modelType']['name'] = 'AssetAdministrationShell'
    dictShell['identification']['idType'] = idType
    dictShell['identification']['id'] = idValue

    dictShell['administration']['version'] = '0'
    dictShell['administration']['revision'] = '1'

    emptyDictionary = {}
    dictShell['hasDataSpecification'].append(emptyDictionary)

    keyAsset = {}
    keyAsset['type'] = refType
    keyAsset['local'] = refLocal
    keyAsset['value'] = state['curAsset']['identification']['id']
    keyAsset['index'] = 0
    keyAsset['idType'] = state['curAsset']['identification']['idType']

    dictShell['asset']['keys'] = []
    dictShell['asset']['keys'].append(keyAsset)

    aasx['assetAdministrationShells'].append(dictShell)

    return dictShell


def add_Asset(state, idShort, idType, idValue, kind):

    dictAsset = {}

    # fill default array or dictionaries (top level)
    dictAsset['hasDataSpecificiation'] = []
    dictAsset['assetIdentificationModelRef'] = {}
    dictAsset['identification'] = {}
    dictAsset['idShort'] = idShort
    dictAsset['modelType'] = {}
    dictAsset['kind'] = kind    # Type or Instance
    dictAsset['descriptions'] = []

    # fill default array or dictionaries (2nd level)
    dictAsset['modelType']['name'] = 'Asset'
    dictAsset['identification']['idType'] = idType
    dictAsset['identification']['id'] = idValue
    dictAsset['assetIdentificationModelRef']['keys'] = []

    aasx['assets'].append(dictAsset)

    return dictAsset 


def add_Submodel(state, idShort, idType, idValue, kind, refType, refLocal, semanticIDType, semanticID):

    dictSubmodel = {}

    # fill default array or dictionaries (top level)
    dictSubmodel['semanticId'] = {}
    dictSubmodel['qualifiers'] = []
    dictSubmodel['hasDataSpecificiation'] = []
    dictSubmodel['identification'] = {}
    dictSubmodel['category'] = ""
    dictSubmodel['idShort'] = idShort
    dictSubmodel['modelType'] = {}
    dictSubmodel['kind'] = kind
    dictSubmodel['submodelElements'] = []

    # fill default array or dictionaries (2nd level)
    dictSubmodel['modelType']['name'] = 'Submodel'
    dictSubmodel['identification']['idType'] = idType
    dictSubmodel['identification']['id'] = idValue
    
    # set options
    if state['options'] != None:
        for option in state['options']:
            if 'category' in option:
                if option['category'].upper() == 'PARAMETER':
                    dictSubmodel['category'] = 'PARAMETER'

                elif option['category'].upper() == 'CONSTANT':
                    dictSubmodel['category'] = 'CONSTANT'

                elif option['category'].upper() == 'VARIABLE':
                    dictSubmodel['category'] = 'VARIABLE'
    

    # fill 'semanticId' of submodel
    dictSubmodel['semanticId']['keys'] = []
    
    if semanticID != None and semanticIDType != None:
        keySubmodel = {}
        keySubmodel['type'] = refType
        keySubmodel['local'] = refLocal
        keySubmodel['value'] = semanticID
        keySubmodel['index'] = 0
        keySubmodel['idType'] = semanticIDType
    
        dictSubmodel['semanticId']['keys'].append(keySubmodel)
    
    # add 'submodels' element in assetAdministrationShells
    keyShellSubmodel = {}
    keyShellSubmodel['keys'] = []
    
    keyShellSubmodelElement = {}
    keyShellSubmodelElement['type'] = 'submodel'
    keyShellSubmodelElement['local'] = 'true'
    keyShellSubmodelElement['value'] = idValue
    keyShellSubmodelElement['index'] = 0
    keyShellSubmodelElement['idType'] = idType
    
    keyShellSubmodel['keys'].append(keyShellSubmodelElement)

    state['curShell']['submodels'].append(keyShellSubmodel)

    
    
    aasx['submodels'].append(dictSubmodel)
    return dictSubmodel 


def add_ConceptDescription(idShort, idType, idValue, prefName, shortName, unit, dataType, definition):

    # find already added concept description
    for cd in  aasx['conceptDescriptions']:
        if (cd['identification']['idType'] == idType) and (cd['identification']['id'] == idValue):
            return cd



    dictDescription = {}

    # fill default array or dictionaries (top level)
    dictDescription['embeddedDataSpecifications'] = []
    dictDescription['identification'] = {}
    dictDescription['idShort'] = idShort
    dictDescription['modelType'] = {}
    dictDescription['isCaseOf'] = []
    dictDescription['descriptions'] = []

    # fill default array or dictionaries (2nd level)
    dictDescription['modelType']['name'] = 'ConceptDescription'
    dictDescription['identification']['idType'] = idType
    dictDescription['identification']['id'] = idValue

    dataSpec = {}
    dataSpec['dataSpecification'] = {}
    dataSpec['dataSpecification']['keys'] = []

    dataSpec['dataSpecificationContent'] = {}
    dataSpec['dataSpecificationContent']['preferredName'] = []
    dataSpec['dataSpecificationContent']['shortName'] = []
   
    if unit == None:
        dataSpec['dataSpecificationContent']['unit'] = ""
    else:
        dataSpec['dataSpecificationContent']['unit'] = unit

    if dataType == None:
        dataSpec['dataSpecificationContent']['dataType'] = ""

    else:
        dataSpec['dataSpecificationContent']['dataType'] = dataType


    dataSpec['dataSpecificationContent']['definition'] = []

    if prefName != None:
        multiText = splitMultiLangText(prefName)
        if len(multiText) > 0:
            dataSpec['dataSpecificationContent']['preferredName'] = multiText

        else:
            outMessage('warning: "preferred-name" of concept description"{}" is invalid'.format(idShort))
    
    
    if shortName != None:
        multiText = splitMultiLangText(shortName)
        if len(multiText) > 0:
            dataSpec['dataSpecificationContent']['shortName'] = multiText

        else:
            outMessage('warning: "short-name" of concept description"{}" is invalid'.format(idShort))


    if definition != None:
        multiText = splitMultiLangText(definition)
        if len(multiText) > 0:
            dataSpec['dataSpecificationContent']['definition'] = multiText

        else:
            outMessage('warning: "definition" of concept description"{}" is invalid'.format(idShort))

    
    dictDescription['embeddedDataSpecifications'].append(dataSpec)
    


    aasx['conceptDescriptions'].append(dictDescription)
    return dictDescription 


def add_MultiLanguageProperty(state, idShort, value, refLocal, conceptDesc):

    dictMLP = {}

    # fill default array or dictionaries (top level)
    dictMLP['value'] = {}
    dictMLP['semanticId'] = {}
    dictMLP['constraints'] = []
    dictMLP['hasDataSpecificiations'] = []
    dictMLP['idShort'] = idShort
    dictMLP['category'] = ""
    dictMLP['modelType'] = {}
    dictMLP['kind'] = 'Instance'

    # fill default array or dictionaries (2nd level)
    dictMLP['modelType']['name'] = 'MultiLanguageProperty'


    # set options
    if state['options'] != None:
        for option in state['options']:
            if 'category' in option:
                if option['category'].upper() == 'PARAMETER':
                    dictMLP['category'] = 'PARAMETER'

                elif option['category'].upper() == 'CONSTANT':
                    dictMLP['category'] = 'CONSTANT'

                elif option['category'].upper() == 'VARIABLE':
                    dictMLP['category'] = 'VARIABLE'
    
            if 'kind' in option:
                if option['kind'].upper() == 'INSTANCE':
                    dictMLP['kind'] = 'Instance'

                elif option['kind'].upper() == 'TEMPLATE':
                    dictMLP['kind'] = 'Template'
 
    # set initial value if it is assigned 
    if value != None:
        multiText = splitMultiLangText(value)
        
        if len(multiText) > 0:
            dictMLP['value']['langString'] = []
            dictMLP['value']['langString'] = multiText 

        else:
            outMessage('warning: "initial value" of multi-language-property"{}" is not valid'.format(idShort))

    else:
        outMessage('warning: "initial value" of multi-language-property"{}" is not assigned'.format(idShort))

    
    # fill 'semanticId'
    dictMLP['semanticId']['keys'] = []

    if conceptDesc != None:
        keyCD = {}
        keyCD['type'] = 'ConceptDescription'
        keyCD['local'] = 'true'
        keyCD['value'] = conceptDesc['identification']['id']
        keyCD['index'] = 0
        keyCD['idType'] = conceptDesc['identification']['idType']
   
        if refLocal != None:
            keyCD['local'] = refLocal

        dictMLP['semanticId']['keys'].append(keyCD)
    
    
    # add to parent submodel element
    if state['parseDepth'] == 'Submodel':
        # this property will be inserted under 'submodel'
        state['curSubmodel']['submodelElements'].append(dictMLP)

    else:
        # this property will be inserted under 'submodel element collection'
        state['curCollection']['value'].append(dictMLP)
     
    return dictMLP 

	

def add_Reference(state, idShort, value, refLocal, conceptDesc):

    dictRef = {}

    # fill default array or dictionaries (top level)
    dictRef['value'] = {}
    dictRef['semanticId'] = {}
    dictRef['constraints'] = []
    dictRef['idShort'] = idShort
    dictRef['modelType'] = {}
    dictRef['kind'] = 'Instance'

    # fill default array or dictionaries (2nd level)
    dictRef['modelType']['name'] = 'ReferenceElement'


    # set options
    if state['options'] != None:
        for option in state['options']:
            if 'kind' in option:
                if option['kind'].upper() == 'INSTANCE':
                    dictRef['kind'] = 'Instance'

                elif option['kind'].upper() == 'TEMPLATE':
                    dictRef['kind'] = 'Template'
  
            if 'category' in option:
                if option['category'].upper() == 'PARAMETER':
                    dictSubmodel['category'] = 'PARAMETER'

                elif option['category'].upper() == 'CONSTANT':
                    dictSubmodel['category'] = 'CONSTANT'

                elif option['category'].upper() == 'VARIABLE':
                    dictSubmodel['category'] = 'VARIABLE'
    

    # set initial value if it is assigned 
    if value != None:
        kvArray = getkeyValueFromString(value)
        
        if len(kvArray) > 0:
            dictRef['value']['keys'] = kvArray.copy()

        else:
            outMessage('warning: "initial value" of Reference"{}" is not valid'.format(idShort))

    else:
        outMessage('warning: "initial value" of Reference"{}" is not assigned'.format(idShort))

    
    # fill 'semanticId'
    dictRef['semanticId']['keys'] = []

    if conceptDesc != None:
        keyCD = {}
        keyCD['type'] = 'ConceptDescription'
        keyCD['local'] = 'true'
        keyCD['value'] = conceptDesc['identification']['id']
        keyCD['index'] = 0
        keyCD['idType'] = conceptDesc['identification']['idType']
   
        if refLocal != None:
            keyCD['local'] = refLocal

        dictRef['semanticId']['keys'].append(keyCD)
    
    
    # add to parent submodel element
    if state['parseDepth'] == 'Submodel':
        # this property will be inserted under 'submodel'
        state['curSubmodel']['submodelElements'].append(dictRef)

    else:
        # this property will be inserted under 'submodel element collection'
        state['curCollection']['value'].append(dictRef)
     
    return dictRef 
	
def add_File(state, idShort, value, refLocal, conceptDesc):

    dictFile = {}

    # fill default array or dictionaries (top level)
    dictFile['mimeType'] = ""
    dictFile['value'] = ""
    dictFile['semanticId'] = {}
    dictFile['constraints'] = []
    dictFile['hasDataSpecificiations'] = []
    dictFile['idShort'] = idShort
    dictFile['category'] = ""
    dictFile['modelType'] = {}
    dictFile['kind'] = 'Instance'

    # fill default array or dictionaries (2nd level)
    dictFile['modelType']['name'] = 'File'


    # set options
    if state['options'] != None:
        for option in state['options']:
            if 'category' in option:
                if option['category'].upper() == 'PARAMETER':
                    dictFile['category'] = 'PARAMETER'

                elif option['category'].upper() == 'CONSTANT':
                    dictFile['category'] = 'CONSTANT'

                elif option['category'].upper() == 'VARIABLE':
                    dictFile['category'] = 'VARIABLE'
    
            if 'kind' in option:
                if option['kind'].upper() == 'INSTANCE':
                    dictFile['kind'] = 'Instance'

                elif option['kind'].upper() == 'TEMPLATE':
                    dictFile['kind'] = 'Template'

            if 'mimetype' in option:
                dictFile['mineType'] = option['mimetype']

 
    # set initial value if it is assigned 
    if value != None:
        dictFile['value'] = value

    else:
        outMessage('warning: "initial value" of File"{}" is not assigned'.format(idShort))

    
    # fill 'semanticId'
    dictFile['semanticId']['keys'] = []

    if conceptDesc != None:
        keyCD = {}
        keyCD['type'] = 'ConceptDescription'
        keyCD['local'] = 'true'
        keyCD['value'] = conceptDesc['identification']['id']
        keyCD['index'] = 0
        keyCD['idType'] = conceptDesc['identification']['idType']
   
        if refLocal != None:
            keyCD['local'] = refLocal

        dictFile['semanticId']['keys'].append(keyCD)
    
    
    # add to parent submodel element
    if state['parseDepth'] == 'Submodel':
        # this property will be inserted under 'submodel'
        state['curSubmodel']['submodelElements'].append(dictFile)

    else:
        # this property will be inserted under 'submodel element collection'
        state['curCollection']['value'].append(dictFile)
     
    return dictFile


def add_SMCollection(state, depth, idShort, refLocal, conceptDesc):

    dictSMC = {}

    # fill default array or dictionaries (top level)
    dictSMC['ordered'] = False
    dictSMC['allowDuplicates'] = False
    dictSMC['semanticId'] = {}
    dictSMC['constraints'] = []
    dictSMC['hasDataSpecificiations'] = []
    dictSMC['idShort'] = idShort
    dictSMC['category'] = ""
    dictSMC['modelType'] = {}
    dictSMC['value'] = []
    dictSMC['kind'] = 'Instance'

    # fill default array or dictionaries (2nd level)
    dictSMC['modelType']['name'] = 'SubmodelElementCollection'

    # set options
    if state['options'] != None:
        for option in state['options']:
            if 'category' in option:
                if option['category'].upper() == 'PARAMETER':
                    dictSMC['category'] = 'PARAMETER'

                elif option['category'].upper() == 'CONSTANT':
                    dictSMC['category'] = 'CONSTANT'

                elif option['category'].upper() == 'VARIABLE':
                    dictSMC['category'] = 'VARIABLE'

            if 'ordered' in option:
                if option['ordered'].lower() == 'true':
                    dictSMC['ordered'] = True

            if 'allowduplicates' in option:
                if option['allowduplicates'].lower() == 'true':
                    dictSMC['allowDuplicates'] = True

            if 'kind' in option:
                if option['kind'].upper() == 'INSTANCE':
                    dictSMC['kind'] = 'Instance'

                elif option['kind'].upper() == 'TEMPLATE':
                    dictSMC['kind'] = 'Template'
 
    # fill 'semanticId'
    dictSMC['semanticId']['keys'] = []
    
    if conceptDesc != None:
        keyCD = {}
        keyCD['type'] = 'ConceptDescription'
        keyCD['local'] = 'true'
        keyCD['value'] = conceptDesc['identification']['id']
        keyCD['index'] = 0
        keyCD['idType'] = conceptDesc['identification']['idType']
   
        if refLocal != None:
            keyCD['local'] = refLocal

        dictSMC['semanticId']['keys'].append(keyCD)
    
    
    # add to parent submodel element
    if state['parseDepth'] == 'Submodel':
        if depth != 0:
            outMessage('error  : invalid SMC depth in SMC "{}"'.format(idShort))
            return None
        
        # this property will be inserted under 'submodel'
        state['curSubmodel']['submodelElements'].append(dictSMC)
        state['curCollection'] = dictSMC

        state['lastSMC'][0] = dictSMC

    else:
        # this property will be inserted under 'submodel element collection'
        #state['curCollection']['value'].append(dictSMC)
        #state['curCollection'] = dictSMC

        if depth == 0:
            state['curSubmodel']['submodelElements'].append(dictSMC)
            state['curCollection'] = dictSMC

            state['lastSMC'][0] = dictSMC

        else:
            state['lastSMC'][depth - 1]['value'].append(dictSMC)
            state['curCollection'] = dictSMC

            state['lastSMC'][depth] = dictSMC
    
    
    return dictSMC 


def add_Property(state, idShort, valueType, value, refLocal, conceptDesc):

    dictProperty = {}

    # fill default array or dictionaries (top level)
    dictProperty['value'] = ""
    dictProperty['semanticId'] = {}
    dictProperty['constraints'] = []
    dictProperty['hasDataSpecificiations'] = []
    dictProperty['idShort'] = idShort
    dictProperty['category'] = ""
    dictProperty['modelType'] = {}
    dictProperty['valueType'] = {}
    dictProperty['kind'] = 'Instance'

    # fill default array or dictionaries (2nd level)
    dictProperty['modelType']['name'] = 'Property'
    dictProperty['valueType']['dataObjectType'] = {}
    if valueType != None:
        dictProperty['valueType']['dataObjectType']['name'] = valueType
    else:
        dictProperty['valueType']['dataObjectType']['name'] = ""
    
    # set options
    if state['options'] != None:
        for option in state['options']:
            if 'category' in option:
                if option['category'].upper() == 'PARAMETER':
                    dictProperty['category'] = 'PARAMETER'

                elif option['category'].upper() == 'CONSTANT':
                    dictProperty['category'] = 'CONSTANT'

                elif option['category'].upper() == 'VARIABLE':
                    dictProperty['category'] = 'VARIABLE'

            if 'kind' in option:
                if option['kind'].upper() == 'INSTANCE':
                    dictProperty['kind'] = 'Instance'

                elif option['kind'].upper() == 'TEMPLATE':
                    dictProperty['kind'] = 'Template'
    

    # set initial value if it is assigned
    if value != None and valueType != None:
        dictProperty['value'] = getValueObject(valueType, value)
    else:
        dictProperty['value'] = ""
        outMessage('warning: "initial value" of property"{}" is not assigned'.format(idShort))

    
    # fill 'semanticId'
    dictProperty['semanticId']['keys'] = []
    
    if conceptDesc != None:
        keyCD = {}
        keyCD['type'] = 'ConceptDescription'
        keyCD['local'] = 'true'
        keyCD['value'] = conceptDesc['identification']['id']
        keyCD['index'] = 0
        keyCD['idType'] = conceptDesc['identification']['idType']
   
        if refLocal != None:
            keyCD['local'] = refLocal

        dictProperty['semanticId']['keys'].append(keyCD)
    
    
    
    # add to parent submodel element
    if state['parseDepth'] == 'Submodel':
        # this property will be inserted under 'submodel'
        state['curSubmodel']['submodelElements'].append(dictProperty)

    else:
        # this property will be inserted under 'submodel element collection'
        state['curCollection']['value'].append(dictProperty)
     

    return dictProperty 








#------------------------ .xlsx Parsing codes ---------------------------------
COLUMN_ASSET                = 'A'
COLUMN_AAS_LEVEL0           = 'B'
COLUMN_AAS_LEVEL1           = 'C'
COLUMN_AAS_LEVEL2           = 'D'
COLUMN_SUBMODEL             = 'E'
COLUMN_COLLECTION_LEVEL0    = 'F'
COLUMN_COLLECTION_LEVEL1    = 'G'
COLUMN_COLLECTION_LEVEL2    = 'H'
COLUMN_COLLECTION_LEVEL3    = 'I'
COLUMN_COLLECTION_LEVEL4    = 'J'
COLUMN_COLLECTION_LEVEL5    = 'K'
COLUMN_FIELD_NAME           = 'L'
COLUMN_PROPERTY             = 'M'
COLUMN_OPTIONS              = 'N'
COLUMN_ASSET_AAS_SM_ID_IRI  = 'O'
COLUMN_REFERENCE_TYPE       = 'P'
COLUMN_REFERENCE_LOCAL      = 'Q'
COLUMN_SEMANTICS_NAME       = 'R'
COLUMN_SEMANTICS_SHORT_NAME = 'S'
COLUMN_SEMANTICS_PREF_NAME  = 'T'
COLUMN_SEMANTICS_DATA_TYPE  = 'U'
COLUMN_SENAMTICS_IRI        = 'V'
COLUMN_SEMANTICS_IRDI       = 'W'
COLUMN_INITIAL_VALUE        = 'X'
COLUMN_ARRAY                = 'Y'
COLUMN_ENGINEERING_UNIT     = 'Z'
COLUMN_PROPERTY_VALUE_TYPE  = 'AA'
COLUMN_SEMANTICS_DEFINITION = 'AB'
COLUMN_FIELD_TAG_NAME       = 'AC'
COLUMN_NOTE                 = 'AD'



def parse_ExcelSheetRow(row, state):
    cellAsset               = None
    cellAAS0                = None
    cellAAS1                = None
    cellAAS2                = None
    cellSubmodel            = None
    cellCollection0         = None
    cellCollection1         = None
    cellCollection2         = None
    cellCollection3         = None
    cellCollection4         = None
    cellCollection5         = None
    cellProperty            = None
    cellOptions             = None
    cellAssetAasSmIdIRI     = None
    cellSemName             = None
    cellSemShortName        = None
    cellSemPrefName         = None
    cellSemDataType         = None
    cellPropertyValueType   = None
    cellReferenceType       = None
    cellReferenceLocal      = None
    cellSemIRI              = None
    cellSemIRDI             = None
    cellFieldName           = None
    cellArray               = None
    cellEngineeringUnit     = None
    cellFieldTagName        = None
    cellSemDef              = None
    cellNote                = None
    cellInitialValue        = None

    state['options'] = None

    # fill local variables from .xlsx cell values   
    for cell in row:
        #print("cell.column = {}, cell.value = {}".format(cell.column, cell.value))
        if cell.value == None:
            continue

        if cell.column == COLUMN_ASSET:
            cellAsset = str(cell.value).strip()
                    
        elif cell.column == COLUMN_AAS_LEVEL0:
            cellAAS0 = str(cell.value).strip()

        elif cell.column == COLUMN_AAS_LEVEL1:
            cellAAS1 = str(cell.value).strip()

        elif cell.column == COLUMN_AAS_LEVEL2:
            cellAAS2 = str(cell.value).strip()

        elif cell.column == COLUMN_SUBMODEL:
            cellSubmodel = str(cell.value).strip()

        elif cell.column == COLUMN_COLLECTION_LEVEL0:
            cellCollection0 = str(cell.value).strip()
        
        elif cell.column == COLUMN_COLLECTION_LEVEL1:
            cellCollection1 = str(cell.value).strip()

        elif cell.column == COLUMN_COLLECTION_LEVEL2:
            cellCollection2 = str(cell.value).strip()
        
        elif cell.column == COLUMN_COLLECTION_LEVEL3:
            cellCollection3 = str(cell.value).strip()
        
        elif cell.column == COLUMN_COLLECTION_LEVEL4:
            cellCollection4 = str(cell.value).strip()
			
        elif cell.column == COLUMN_COLLECTION_LEVEL5:
            cellCollection5 = str(cell.value).strip()

        elif cell.column == COLUMN_PROPERTY:
            cellProperty = str(cell.value).strip()

        elif cell.column == COLUMN_OPTIONS:
            cellOptions = str(cell.value).strip()
            if len(cellOptions) > 0:
                state['options'] = splitOptionsText(cellOptions)
                if len(state['options']) == 0:
                    state['options'] = None
            else:
                state['options'] = None

        elif cell.column == COLUMN_ASSET_AAS_SM_ID_IRI:
            cellAssetAasSmIdIRI = str(cell.value).strip()

        elif cell.column == COLUMN_SEMANTICS_NAME:
            cellSemName = str(cell.value).strip()

        elif cell.column == COLUMN_REFERENCE_TYPE:
            cellReferenceType = str(cell.value).strip()

            if cellReferenceType.lower() == 'asset':
                cellReferenceType = 'Asset'
            elif cellReferenceType.lower() == 'globalreference':
                cellReferenceType = 'GlobalReference'
            elif cellReferenceType.lower() == 'submodel':
                cellReferenceType = 'Submodel'
            elif cellReferenceType.lower() == 'conceptdescription':
                cellReferenceType = 'ConceptDescription'
            else:
                outMessage('reference-type "{}" is not supported'.format(cellReferenceType))
                cellReferenceType = None

        elif cell.column == COLUMN_REFERENCE_LOCAL:
            cellReferenceLocal = str(cell.value).strip()

            if cellReferenceLocal.lower() == 'true':
                cellReferenceLocal = 'true'
            elif cellReferenceLocal.lower() == 'false':
                cellReferenceLocal = 'false'
            else:
                outMessage('reference-local "{}" is not supported'.format(cellReferenceLocal))
                cellReferenceLocal = None

        elif cell.column == COLUMN_SENAMTICS_IRI:
            cellSemIRI = str(cell.value).strip()

        elif cell.column == COLUMN_SEMANTICS_IRDI:
            cellSemIRDI = str(cell.value).strip()

        elif cell.column == COLUMN_FIELD_NAME:
            cellFieldName = str(cell.value).strip()

        elif cell.column == COLUMN_ARRAY:
            cellArray = str(cell.value).strip()

        elif cell.column == COLUMN_ENGINEERING_UNIT:
            cellEngineeringUnit = str(cell.value).strip()

        elif cell.column == COLUMN_FIELD_TAG_NAME:
            cellFieldTagName = str(cell.value).strip()

        elif cell.column == COLUMN_SEMANTICS_DEFINITION:
            cellSemDef = str(cell.value).strip()
            if cellSemDef[0] == "'":
                cellSemDef = cellSemDef[1:]

        elif cell.column == COLUMN_NOTE:
            cellNote = str(cell.value).strip()

        elif cell.column == COLUMN_SEMANTICS_SHORT_NAME:
            cellSemShortName = str(cell.value).strip()
            if cellSemShortName[0] == "'":
                cellSemShortName = cellSemShortName[1:]

        elif cell.column == COLUMN_SEMANTICS_PREF_NAME:
            cellSemPrefName = str(cell.value).strip()
            if cellSemPrefName[0] == "'":
                cellSemPrefName = cellSemPrefName[1:]

        elif cell.column == COLUMN_SEMANTICS_DATA_TYPE:
            cellSemDataType = str(cell.value).strip()

        elif cell.column == COLUMN_PROPERTY_VALUE_TYPE:
            cellPropertyValueType = str(cell.value).strip()
            cellPropertyValueType = getValueTypeText(cellPropertyValueType)


        elif cell.column == COLUMN_INITIAL_VALUE:
            cellInitialValue = str(cell.value).strip()
    

    #if state['options'] != None:
    #    print(state['options'])


    # add 'Asset'
    if cellAsset != None:
        if cellAssetAasSmIdIRI != None:
            state['curAsset'] = add_Asset(state, cellAsset, 'IRI', cellAssetAasSmIdIRI, 'Instance')
            state['parseDepth'] = 'Asset'

        else:
            outMessage('error  : Asset "{}" should have "asset/aas/submodel id (IRI)"'.format(cellAsset))
        
        return state


    # add 'AAS'
    cellAAS = None

    if cellAAS0 != None:
        cellAAS = cellAAS0

    elif cellAAS1 != None:
        cellAAS = cellAAS1

    elif cellAAS2 != None:
        cellAAS = cellAAS2


    if cellAAS != None:
        if state['parseDepth'] != 'Asset' :
            outMessage('error  : Asset have to be defined before AAS "{}" is defined'.format(cellAAS))
            return state

        if (cellReferenceType == None) or (cellReferenceLocal == None):
            outMessage('error  : "reference-type" or "reference-local" must have value to define AAS "{}"'.format(cellAAS))
            return state

        if cellAssetAasSmIdIRI != None:
            state['curShell'] = add_Shell(state, cellAAS, 'IRI', cellAssetAasSmIdIRI, cellReferenceType, cellReferenceLocal)
            state['parseDepth'] = 'Shell'

        else:
            outMessage('error  : AAS "{}" should have "asset/aas/submodel id (IRI)"'.format(cellAAS))
        
        return state



    # add 'Submodel' 
    if cellSubmodel != None: 
        if state['parseDepth'] == 'Asset':
            outMessage('error  : Submodel "{}" have to be defined only after AAS is already defined'.format(cellSubmodel))
            return state
        
        if (cellReferenceType == None) or (cellReferenceLocal == None):
            outMessage('error  : "reference-type" or "reference-local" must have value to define Submodel "{}"'.format(cellSubmodel))
            return state

        # semantic-ID changed from Mandatory to Optional        
        #if (cellSemIRI == None) and (cellSemIRDI  == None):
        #    outMessage('"semantic-id-IRDI" or "semantic-id-IRI" must have value to define Submodel "{}"'.format(cellSubmodel))
        #    return state
    
        if cellAssetAasSmIdIRI != None:
            if cellSemIRI != None:
                state['curSubmodel'] = add_Submodel(state, cellSubmodel, 'IRI', cellAssetAasSmIdIRI, 'Instance', cellReferenceType, cellReferenceLocal, 'IRI', cellSemIRI)

            else:
                state['curSubmodel'] = add_Submodel(state, cellSubmodel, 'IRI', cellAssetAasSmIdIRI, 'Instance', cellReferenceType, cellReferenceLocal, 'IRDI', cellSemIRDI)

            state['curCollection'] = None
            state['parseDepth'] = 'Submodel'

            for lastSMC in state['lastSMC']:
                lastSMC = None

        else:
            outMessage('error  : Submodel "{}" should have "asset/aas/submodel id (IRI)"'.format(cellSubmodel))
 
        return state


    # add 'Property' or 'MultiLanguageProperty' with 'ConceptDescription'
    if cellProperty != None:
        smeType, smeIdShort = cellProperty.split(':')

        if (smeType == None) or (smeIdShort == None):
            outMessage('warning: Property"{}" has invalid value (assumes "Prop")'.format(cellProperty))
            smeType = 'Prop'
            smeIdShort = cellProperty

        else:
            smeType = smeType.strip()
            smeIdShort = smeIdShort.strip()


        # semanticID changed from Mandatory to Optional
        if (cellSemIRI == None) and (cellSemIRDI  == None):
            outMessage('warning: "semantic-id-IRDI" or "semantic-id-IRI" must have value to define Property "{}"'.format(smeIdShort))
            #return state
        else:
            if cellSemName == None:
                outMessage('warning: "semantic-name" must have value to define Property "{}"'.format(smeIdShort))
                #return state
        
            if cellSemShortName == None:
                outMessage('warning: "semantic-short-name" must have value to define Property "{}"'.format(smeIdShort))
                #return state
        
            if cellSemPrefName == None:
                outMessage('warning: "semantic-preferred-name" must have value to define Property "{}"'.format(smeIdShort))
                #return state
        
            #if cellSemanticsDataType == None:
            #    outMessage('"semantic-datatype" must have value to define Property "{}"'.format(cellProperty))
            #    return state
        
            if cellPropertyValueType == None and (smeType.lower() == 'mlp' or smeType.lower() == 'prop'):
                outMessage('warning: "property-value-type" must have value to define Property "{}"'.format(smeIdShort))
                #return state
        
            if cellSemDef == None:
                outMessage('warning: it is recommanded to define "semantic-definition" to define Property "{}"'.format(smeIdShort))
                #return state
        
        if (state['parseDepth'] != 'Submodel') and (state['parseDepth'] != 'Collection'):
            outMessage('error: Property "{}" have to be defined only after Submodel or SubmodelCollection is already defined'.format(smeIdShort))
            return state

        cd = None

        if cellSemIRDI != None:
            cd = add_ConceptDescription(cellSemName, 'IRDI', cellSemIRDI, cellSemPrefName, cellSemShortName, cellEngineeringUnit, cellSemDataType, cellSemDef)
        
        elif cellSemIRI != None:
            cd = add_ConceptDescription(cellSemName, 'IRI', cellSemIRI, cellSemPrefName, cellSemShortName, cellEngineeringUnit, cellSemDataType, cellSemDef)
            
        if smeType.lower() == 'mlp':
            add_MultiLanguageProperty(state, smeIdShort, cellInitialValue, cellReferenceLocal, cd)

        elif smeType.lower() == 'prop':
            add_Property(state, smeIdShort, cellPropertyValueType, cellInitialValue, cellReferenceLocal, cd)

        elif smeType.lower() == 'file':
            add_File(state, smeIdShort, cellInitialValue, cellReferenceLocal, cd)
		
        elif smeType.lower() == 'ref':
            add_Reference(state, smeIdShort, cellInitialValue, cellReferenceLocal, cd)
			
        else:
            outMessage('warning: "unknown submodel element type:{}" fail to add Property "{}"'.format(smeType, smeIdShort))


        return state

    # add 'SubmodelElementCollection'
    cellSMC = None
    smcDepth = 0

    if cellCollection0 != None:
        cellSMC = cellCollection0
        smcDepth = 0

    elif cellCollection1 != None:
        cellSMC = cellCollection1
        smcDepth = 1

    elif cellCollection2 != None:
        cellSMC = cellCollection2
        smcDepth = 2

    elif cellCollection3 != None:
        cellSMC = cellCollection3
        smcDepth = 3

    elif cellCollection4 != None:
        cellSMC = cellCollection4
        smcDepth = 4
		
    elif cellCollection5 != None:
        cellSMC = cellCollection5
        smcDepth = 5


    if cellSMC != None:
      
        # semanticId changed from Mandatory to Optional
        if (cellSemIRI == None) and (cellSemIRDI  == None):
            outMessage('warning: "semantic-id-IRDI" or "semantic-id-IRI" must have value to define submodel-element-collection "{}"'.format(cellSMC))
            #return state
        else:
            if cellSemName == None:
                outMessage('warning: "semantic-name" must have value to define submodel-element-collection "{}"'.format(cellSMC))
                #return state
        
            if cellSemShortName == None:
                outMessage('warning: "semantic-short-name" must have value to define submodel-element-collection "{}"'.format(cellSMC))
                #return state
        
            if cellSemPrefName == None:
                outMessage('warning: "semantic-preferred-name" must have value to define submodel-element-collection "{}"'.format(cellSMC))
                #return state
        
            if cellSemDef == None:
                outMessage('warning: "semantic-definition" must have value to define submodel-element-collection "{}"'.format(cellSMC))
                #return state
        
        if (state['parseDepth'] != 'Submodel') and (state['parseDepth'] != 'Collection'):
            outMessage('error  : submodel-element-collection "{}" have to be defined only after Submodel or SubmodelCollection is already defined'.format(cellSMC))
            return state


        cd = None

        if cellSemIRDI != None:
            cd = add_ConceptDescription(cellSemName, 'IRDI', cellSemIRDI, cellSemPrefName, cellSemShortName, cellEngineeringUnit, cellSemDataType, cellSemDef)

        elif cellSemIRI != None:
            cd = add_ConceptDescription(cellSemName, 'IRI', cellSemIRI, cellSemPrefName, cellSemShortName, cellEngineeringUnit, cellSemDataType, cellSemDef)

        
        add_SMCollection(state, smcDepth, cellSMC, cellReferenceLocal, cd)
        state['parseDepth'] = 'Collection'
        return state

    return state


def parse_ExcelSheet(sheet):

    parseState = {}
    parseState['parseDepth'] = 'idle'
    parseState['curAsset'] = None
    parseState['curShell'] = None
    parseState['curSubmodel'] = None
    parseState['curCollection'] = None
    parseState['options'] = None
    parseState['lastSMC'] = [None, None, None, None, None]

    # add aas-objects from .xlsx file
    skipRow = 2
    for row in sheet.rows:
        if skipRow > 0:
            skipRow -= 1
            continue

        parseState = parse_ExcelSheetRow(row, parseState)

    # make relations between added aas-objects

    return


#initialize AAS.json dictionaries
del aasx['assetAdministrationShells'][:]
del aasx['assets'][:]
del aasx['submodels'][:]
del aasx['conceptDescriptions'][:]

#outMessgae('Parser version: {}'.format(PARSER_VERSION_STRING))
outMessage('Parser version: {}'.format(PARSER_VERSION_STRING))


# open input .xlsx file
xls_wb          = xl.load_workbook(sys.argv[1])
xls_sheetNames  = xls_wb.get_sheet_names()


# parsing every sheets in .xlsx file
for sheetName in xls_sheetNames:
    sheet = xls_wb[ sheetName ]
    parse_ExcelSheet(sheet)

xls_wb.close()
outMessage('completed')

# close result file
fp_result.close()

# make result aasx.json file
#chohpower
with open(sys.argv[2], 'w') as fp:
    json.dump(aasx, fp, ensure_ascii = False, indent = 4)


