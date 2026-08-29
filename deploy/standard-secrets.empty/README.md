# Empty managed-source secret directory

This tracked empty directory makes upload-only deployments work without a managed HTTP secret. For managed HTTP sources, set `STANDARD_SECRETS_DIR` to a protected host directory containing one file per `auth_secret_ref`; do not place secret values in this repository.
