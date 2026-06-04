"""Third-party analytics integration for the hosted Streamlit app."""

from __future__ import annotations

import streamlit.components.v1 as components

CLARITY_PROJECT_ID = "x1tq0msm2n"


def render_clarity_analytics() -> None:
    """Load Microsoft Clarity without adding visible interface content."""
    components.html(
        f"""
        <script type="text/javascript">
            (function(c,l,a,r,i,t,y){{
                c[a]=c[a]||function(){{(c[a].q=c[a].q||[]).push(arguments)}};
                t=l.createElement(r);t.async=1;t.src="https://www.clarity.ms/tag/"+i;
                y=l.getElementsByTagName(r)[0];y.parentNode.insertBefore(t,y);
            }})(window, document, "clarity", "script", "{CLARITY_PROJECT_ID}");
        </script>
        """,
        height=0,
        width=0,
    )
