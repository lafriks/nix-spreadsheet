dnl Warning: This is an automatically generated file, do not edit!
dnl Process this file with autoconf to produce a configure script.
AC_PREREQ([2.54])
AC_INIT([], [0.1])
AM_INIT_AUTOMAKE([foreign])

dnl pkg-config
AC_PATH_PROG(PKG_CONFIG, pkg-config, no)
if test "x$PKG_CONFIG" = "xno"; then
        AC_MSG_ERROR([You need to install pkg-config])
fi

AC_PROG_INSTALL

AC_PATH_PROG(MCS, mcs, no)
if test "x$MCS" = "xno"; then
        AC_MSG_ERROR([mcs Not found])
fi




AC_ARG_ENABLE(debug,
	AC_HELP_STRING([--enable-debug],
		[Use 'Debug' Configuration [default=NO]]),
		enable_debug=yes, enable_debug=no)
AM_CONDITIONAL(ENABLE_DEBUG, test x$enable_debug = xyes)
if test "x$enable_debug" = "xyes" ; then
	COMPOUNDFILE_DEBUG_LIB='Nix.CompoundFile.dll'
	AC_SUBST(COMPOUNDFILE_DEBUG_LIB)
	SPREADSHEET_DEBUG_LIB='Nix.ExcelWriter.dll'
	AC_SUBST(SPREADSHEET_DEBUG_LIB)
	DEBUG_CONFIG_LIBRARIES=' ${pkglibdir}/Nix.CompoundFile.dll ${pkglibdir}/Nix.ExcelWriter.dll'
	DEBUG_CONFIG_LIBS=' -r:${pkglibdir}/Nix.CompoundFile.dll -r:${pkglibdir}/Nix.ExcelWriter.dll'
	AC_SUBST(DEBUG_CONFIG_LIBRARIES)
	AC_SUBST(DEBUG_CONFIG_LIBS)
	CONFIG_REQUESTED="yes"
fi
AC_ARG_ENABLE(release,
	AC_HELP_STRING([--enable-release],
		[Use 'Release' Configuration [default=YES]]),
		enable_release=yes, enable_release=no)
AM_CONDITIONAL(ENABLE_RELEASE, test x$enable_release = xyes)
if test "x$enable_release" = "xyes" ; then
	COMPOUNDFILE_RELEASE_LIB='Nix.CompoundFile.dll'
	AC_SUBST(COMPOUNDFILE_RELEASE_LIB)
	SPREADSHEET_RELEASE_LIB='Nix.ExcelWriter.dll'
	AC_SUBST(SPREADSHEET_RELEASE_LIB)
	RELEASE_CONFIG_LIBRARIES=' ${pkglibdir}/Nix.CompoundFile.dll ${pkglibdir}/Nix.ExcelWriter.dll'
	RELEASE_CONFIG_LIBS=' -r:${pkglibdir}/Nix.CompoundFile.dll -r:${pkglibdir}/Nix.ExcelWriter.dll'
	AC_SUBST(RELEASE_CONFIG_LIBRARIES)
	AC_SUBST(RELEASE_CONFIG_LIBS)
	CONFIG_REQUESTED="yes"
fi
if test -z "$CONFIG_REQUESTED" ; then
	COMPOUNDFILE_RELEASE_LIB='Nix.CompoundFile.dll'
	AC_SUBST(COMPOUNDFILE_RELEASE_LIB)
	SPREADSHEET_RELEASE_LIB='Nix.ExcelWriter.dll'
	AC_SUBST(SPREADSHEET_RELEASE_LIB)
	RELEASE_CONFIG_LIBRARIES=' ${pkglibdir}/Nix.CompoundFile.dll ${pkglibdir}/Nix.ExcelWriter.dll'
	RELEASE_CONFIG_LIBS=' -r:${pkglibdir}/Nix.CompoundFile.dll -r:${pkglibdir}/Nix.ExcelWriter.dll'
	AC_SUBST(RELEASE_CONFIG_LIBRARIES)
	AC_SUBST(RELEASE_CONFIG_LIBS)
	AM_CONDITIONAL(ENABLE_RELEASE, true)
fi


AC_CONFIG_FILES([
./Test/test
./ExcelWriter/Makefile
./Test/Makefile
./Makefile
./ExcelWriter/.SpreadSheet.pc
./Test/.Test.pc
./CompoundFile/.CompoundFile.pc
./CompoundFile/Makefile

])

AC_OUTPUT