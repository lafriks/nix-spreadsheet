
# Warning: This is an automatically generated file, do not edit!

if ENABLE_DEBUG
ASSEMBLY_COMPILER_COMMAND = mcs
ASSEMBLY_COMPILER_FLAGS =  -noconfig -codepage:utf8 -warn:4 -debug -d:DEBUG
ASSEMBLY = ./bin/Debug/Nix.SpreadSheet.dll
COMPILE_TARGET = library
PC_FILES_IN = .SpreadSheet.pc.in
PROJECT_REFERENCES =  \
	../CompoundFile/bin/Debug/Nix.CompoundFile.dll
BUILD_DIR = ./bin/Debug/
endif
if ENABLE_RELEASE
ASSEMBLY_COMPILER_COMMAND = mcs
ASSEMBLY_COMPILER_FLAGS =  -noconfig -codepage:utf8 -warn:4
ASSEMBLY = ./bin/Release/Nix.SpreadSheet.dll
COMPILE_TARGET = library
PC_FILES_IN = .SpreadSheet.pc.in
PROJECT_REFERENCES =  \
	../CompoundFile/bin/Release/Nix.CompoundFile.dll
BUILD_DIR = ./bin/Release/
endif

	
all: $(ASSEMBLY)

FILES = \
	./AssemblyInfo.cs \
	./AbstractCell.cs \
	./Cell.cs \
	./Comment.cs \
	./Font.cs \
	./SpreadSheetDocument.cs \
	./SpreadSheetFileFormat.cs \
	./StringFormating.cs \
	./Sheet.cs \
	./Style.cs \
	./Utils.cs \
	./Provider/IFileProvider.cs \
	./Provider/XlsProvider.cs \
	./Provider/BIFF/BIFFRecord.cs \
	./Provider/BIFF/BIFFStringHelper.cs \
	./Provider/BIFF/EOF.cs \
	./Provider/BIFF/FONT.cs \
	./Provider/BIFF/BIFF8/BOF.cs

DATA_FILES = 

RESOURCES = 

EXTRAS = 

REFERENCES =  \
	-r:System \
	-r:System.Data \
	-r:System.Xml

DLL_REFERENCES = 

$(ASSEMBLY) $(ASSEMBLY).mdb: $(build_sources) $(build_resources) $(build_datafiles) 
	mkdir -p $(dir $(ASSEMBLY))		

	$(ASSEMBLY_COMPILER_COMMAND) $(ASSEMBLY_COMPILER_FLAGS) -out:$@ -target:$(COMPILE_TARGET) $(build_sources) $(build_resources_embed) $(build_references_ref)

include $(top_srcdir)/Makefile.include
