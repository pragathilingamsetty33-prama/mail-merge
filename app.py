import io
import tempfile
import zipfile
from pathlib import Path

import streamlit as st

from mail_merge import CertificatePipeline


BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "templates" / "template_config.json"

st.set_page_config(
    page_title="Certificate Mail Merge",
    page_icon="🎓",
    layout="centered",
)

st.title("🎓 Certificate Mail Merge")
st.write(
    "Upload a CSV or Excel recipient list and generate personalized certificate "
    "PDF and PNG files."
)

with st.expander("Required spreadsheet columns"):
    st.markdown(
        "- **Required:** `name`, `course`\n"
        "- **Optional:** `date`, `cert_id`, `instructor`\n\n"
        "Column names such as `Full Name`, `Course Name`, and `Certificate ID` "
        "are also supported."
    )

uploaded_file = st.file_uploader(
    "Upload recipient data",
    type=["csv", "xlsx", "xls"],
    help="The file must contain at least name and course columns.",
)
export_format = st.selectbox(
    "Output format",
    options=["both", "pdf", "png"],
    format_func=lambda value: {
        "both": "PDF + PNG",
        "pdf": "PDF only",
        "png": "PNG only",
    }[value],
)

generate = st.button("Generate certificates", type="primary", disabled=uploaded_file is None)

if generate and uploaded_file is not None:
    suffix = Path(uploaded_file.name).suffix.lower()

    with tempfile.TemporaryDirectory() as temporary_directory:
        temp_dir = Path(temporary_directory)
        data_path = temp_dir / f"recipients{suffix}"
        output_dir = temp_dir / "output_certificates"
        log_path = temp_dir / "generation.log"
        data_path.write_bytes(uploaded_file.getvalue())

        try:
            with st.spinner("Generating certificates..."):
                pipeline = CertificatePipeline(
                    data_file=str(data_path),
                    template_config=str(TEMPLATE_PATH),
                    output_dir=str(output_dir),
                    export_format=export_format,
                    dpi=300,
                    log_file=str(log_path),
                )
                results = pipeline.run()

            st.success(
                f"Completed: {results['successful']} generated, "
                f"{results['skipped']} skipped, {results['failed']} failed."
            )

            generated_files = [
                path for path in output_dir.rglob("*") if path.is_file()
            ]

            if not generated_files:
                st.warning("No certificates were generated. Check that the file has valid name and course values.")
            else:
                archive = io.BytesIO()
                with zipfile.ZipFile(archive, "w", zipfile.ZIP_DEFLATED) as zip_file:
                    for file_path in generated_files:
                        zip_file.write(file_path, file_path.relative_to(output_dir))
                archive.seek(0)

                st.download_button(
                    "Download generated certificates (ZIP)",
                    data=archive.getvalue(),
                    file_name="generated_certificates.zip",
                    mime="application/zip",
                    type="primary",
                )

                with st.expander("Generated files"):
                    st.write([str(path.relative_to(output_dir)) for path in generated_files])

        except Exception as error:
            st.error(f"Generation failed: {error}")
            st.exception(error)

st.divider()
st.caption("Files are processed temporarily and are not stored after the session ends.")
