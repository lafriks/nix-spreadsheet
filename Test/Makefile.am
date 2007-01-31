
# Warning: This is an automatically generated file, do not edit!

if ENABLE_DEBUG
ASSEMBLY_COMPILER_COMMAND = mcs
ASSEMBLY_COMPILER_FLAGS =  -noconfig -codepage:utf8 -warn:4 -debug -d:DEBUG "-main:Test.Test"

ASSEMBLY = ./bin/Debug/Test.exe
COMPILE_TARGET = exe
ASSEMBLY_WRAPPER = test
ASSEMBLY_WRAPPER_IN = test.in
PROJECT_REFERENCES =  \
	../ExcelWriter/bin/Debug/Nix.ExcelWriter.dll
BUILD_DIR = ./bin/Debug/
endif
if ENABLE_RELEASE
ASSEMBLY_COMPILER_COMMAND = mcs
ASSEMBLY_COMPILER_FLAGS =  -noconfig -codepage:utf8 -warn:4 "-main:Test.Test"

ASSEMBLY = ./bin/Release/Test.exe
COMPILE_TARGET = exe
ASSEMBLY_WRAPPER = test
ASSEMBLY_WRAPPER_IN = test.in
PROJECT_REFERENCES =  \
	../ExcelWriter/bin/Release/Nix.ExcelWriter.dll
BUILD_DIR = ./bin/Release/
endif

	
all: $(ASSEMBLY)

FILES = \
	./AssemblyInfo.cs \
	./Test.cs 

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
