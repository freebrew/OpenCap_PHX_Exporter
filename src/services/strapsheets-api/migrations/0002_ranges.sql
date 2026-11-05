-- Measure ranges per family step. Mirrored to FormOptions!F:I when a sheet is configured.
CREATE TABLE ranges (
  family TEXT NOT NULL,
  step_id TEXT NOT NULL,
  min REAL,
  max REAL,
  PRIMARY KEY (family, step_id)
);
