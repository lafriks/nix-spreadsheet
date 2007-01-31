
# Warning: This is an automatically generated file, do not edit!

if ENABLE_DEBUG
ASSEMBLY_COMPILER_COMMAND = mcs
ASSEMBLY_COMPILER_FLAGS =  -noconfig -codepage:utf8 -warn:4 -debug -d:DEBUG
ASSEMBLY = ./bin/Debug/Nix.CompoundFile.dll
COMPILE_TARGET = library
PC_FILES_IN = .CompoundFile.pc.in
PROJECT_REFERENCES = 
BUILD_DIR = ./bin/Debug/
endif
if ENABLE_RELEASE
ASSEMBLY_COMPILER_COMMAND = mcs
ASSEMBLY_COMPILER_FLAGS =  -noconfig -codepage:utf8 -warn:4
ASSEMBLY = ./bin/Release/Nix.CompoundFile.dll
COMPILE_TARGET = library
PC_FILES_IN = .CompoundFile.pc.in
PROJECT_REFERENCES = 
BUILD_DIR = ./bin/Release/
endif

	
all: $(ASSEMBLY)

FILES = \
	./AssemblyInfo.cs \
	./Common.cs \
	./EndianWriter.cs \
	./LittleEndianWriter.cs \
	./Ole2CompoundFile.cs \
	./Ole2DirectoryEntry.cs \
	./Ole2Storage.cs \
	./Ole2Stream.cs \
	./RedBlackEnumerator.cs \
	./RedBlackTree.cs \
	./Managers/DirectoryManager.cs \
	./Managers/MasterSectorAllocationManager.cs \
	./Managers/SectorAllocationManager.cs 

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
