We use the following architecture for VBA projects. This defines how the user interface links to driver subroutines that instance classes and call their procedure methods. We use xlwings vba edit in VS Code, so the Project folder's `src` folder contains code files for each module in addition to the project's \*.xlsm file. Legacy projects may contain non-conforming code structures. We refactor to conform to these architectures as appropriate. 

**Project Code**
```
ProjectName  (VBAProject)
├── Constants   (Constants.bas)   
├── Interface   (Interface.bas)
├── ClassXYZ   (ClassXYZ.cls)
│   ├── TopLevelProcedure
│   │   ├── Method1
│   │   ├── Method2
│   │   ├── SubProcedure1
│   │   │   ├── SubMethod1
│   │   │   ├── SubMethod2
│   │   │   └── etc.
│   │   ├── Method3
│   │   └── etc.
├── Utilities   (Utilities.bas)
└── Validation   (Validation.bas)   
```
* Procedures (`TopLevelProcedure`, `SubProcedure1` etc. call single-action methods to execute a larger task. Sub-procedures are warranted if a top-level task needs to be broken into multiple steps and for cases where a multistep task needs to execute multiple times within a top-level procedure
* `Constants`: Global constants used within project code
* `Interface`: Top-level driver subroutines user-initiated (by buttons, menu commands etc.) that toggle Application attributes for optimizing performance, call class procedures, and take care of top-level error reporting if errors occur in procedures
* `Utilities`: Generic, utility subs and functions
* `Validation`: Contains factory functions to instance project classes for testing (called by test suite). This is needed since an external workbook cannot instance project classes in VBA

**Test Code**
```
tests_ProjectName  (VBAProject)
├── Populate   (Populate.bas)
├── tests_UseCase1   (tests_UseCase1.bas)
├── tests_UseCase2   (tests_UseCase2.bas)
├── tests_UseCaseX   (tests_UseCase3.bas)
├── Procedures   (Procedures.cls)
├── Procedure   (Procedure.cls)
├── Utilities   (Utilities.bas)
└── Test   (Test.cls)
```

`Populate`: Module to generate and populate templates for testing including importing needed test data
`tests_UseCaseX` modules: Modules containing tests for one or more Procedures (`Procedure` attributes instanced as `proc` with hard-coded definitions in `Procedures` class. See `create_new_test_procedure.md` skill is a reference on creating new tests_UseCaseX modules and `Procedure` attributes to group individual tests logically by use case)