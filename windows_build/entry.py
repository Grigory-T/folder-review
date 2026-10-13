"""Windows executable entry point; the scanner itself is unmodified."""

import folder_review
from windows_build.bootstrap import configure_frozen_application

configure_frozen_application(folder_review)
raise SystemExit(folder_review.main())
