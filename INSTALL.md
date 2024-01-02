Do once if required:
   * cpanm -n Dist::Zilla
   * dzil setup 

Install prerequisites:
   * dzil authordeps --missing | cpanm -n 
   * dzil listdeps --author --missing | cpanm -n

Test:
   * dzil test --all

Build:
   * dzil build

Upload new release (for maintainers only):
   * dzil release

