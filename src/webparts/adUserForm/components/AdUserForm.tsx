import * as React from "react";
import {
  DefaultButton,
  Dropdown,
  Icon,
  IconButton,
  IDropdownOption,
  MessageBar,
  MessageBarType,
  Persona,
  PersonaSize,
  PrimaryButton,
  ProgressIndicator,
  Spinner,
  SpinnerSize,
  TextField,
  Toggle,
} from "@fluentui/react";
import styles from "./AdUserForm.module.scss";
import { IAdUserFormProps } from "./IAdUserFormProps";

interface IFormData {
  firstName: string;
  lastName: string;
  company: string;
  jobTitle: string;
  username: string;
  password: string;
  confirmPassword: string;
  forcePasswordChange: boolean;
  businessPhone: string;
  mobilePhone: string;
  fax: string;
  address: string;
  city: string;
  license: string;
  photo: File | undefined;
}

type PasswordTone = "none" | "weak" | "medium" | "strong";

const createInitialFormData = (): IFormData => ({
  firstName: "",
  lastName: "",
  company: "",
  jobTitle: "",
  username: "",
  password: "",
  confirmPassword: "",
  forcePasswordChange: true,
  businessPhone: "",
  mobilePhone: "",
  fax: "",
  address: "",
  city: "",
  license: "",
  photo: undefined,
});

const friendlyNameMap: Record<string, string> = {
  ENTERPRISEPACK: "Microsoft 365 E3",
  ENTERPRISEPREMIUM: "Microsoft 365 E5",
  DEVELOPERPACK_E5: "Microsoft 365 E5 Developer",
  BUSINESS_PREMIUM: "Microsoft 365 Business Premium",
  STANDARDPACK: "Office 365 E1",
  DESKLESSPACK: "Microsoft 365 F1",
  M365_F3: "Microsoft 365 F3",
  FLOW_FREE: "Power Automate Free",
  POWERAPPS_VIRAL: "Power Apps Free",
  PROJECT_P1: "Project Plan 1",
  PROJECTPROFESSIONAL: "Project Professional",
  VISIOCLIENT: "Visio Plan 2",
};

const validatePassword = (password: string): { label: string; tone: PasswordTone } => {
  const strong = /^(?=.*[a-z])(?=.*[A-Z])(?=.*\d)(?=.*[^A-Za-z\d]).{8,}$/;
  const medium = /^(?=.*[A-Z])(?=.*\d).{6,}$/;

  if (!password) {
    return { label: "", tone: "none" };
  }

  if (strong.test(password)) {
    return { label: "Strong password", tone: "strong" };
  }

  if (medium.test(password)) {
    return { label: "Moderate password", tone: "medium" };
  }

  return { label: "Weak password", tone: "weak" };
};

const renderSectionHeading = (iconName: string, eyebrow: string, title: string, description: string): JSX.Element => (
  <div className={styles.sectionHeading}>
    <div className={styles.sectionIcon}>
      <Icon iconName={iconName} />
    </div>
    <div>
      <div className={styles.sectionEyebrow}>{eyebrow}</div>
      <h3 className={styles.sectionTitle}>{title}</h3>
      <p className={styles.sectionDescription}>{description}</p>
    </div>
  </div>
);

const AdUserForm: React.FC<IAdUserFormProps> = ({ graphClient, onClose }) => {
  const [formData, setFormData] = React.useState<IFormData>(createInitialFormData);
  const [photoPreview, setPhotoPreview] = React.useState<string | null>(null);
  const [licenseOptions, setLicenseOptions] = React.useState<IDropdownOption[]>([]);
  const [verifiedDomain, setVerifiedDomain] = React.useState("sj30b.onmicrosoft.com");
  const [loadingTenantData, setLoadingTenantData] = React.useState(true);
  const [loading, setLoading] = React.useState(false);
  const [error, setError] = React.useState<string | null>(null);
  const [successMessage, setSuccessMessage] = React.useState<string | null>(null);
  const [passwordVisible, setPasswordVisible] = React.useState(false);
  const [passwordStrength, setPasswordStrength] = React.useState<{ label: string; tone: PasswordTone }>({
    label: "",
    tone: "none",
  });
  const [usernameAvailable, setUsernameAvailable] = React.useState<boolean | null>(null);
  const fileInputRef = React.useRef<HTMLInputElement>(null);
  const suggestedUsernameRef = React.useRef("");

  React.useEffect(() => {
    let isMounted = true;

    const fetchTenantData = async (): Promise<void> => {
      setLoadingTenantData(true);

      try {
        const orgResponse = await graphClient.api("/organization").get();
        const domain =
          orgResponse.value?.[0]?.verifiedDomains?.find((item: { isDefault?: boolean; name?: string }) => item.isDefault)?.name ??
          "sj30b.onmicrosoft.com";

        const licenseResponse = await graphClient.api("/subscribedSkus").get();
        const activeLicenses = (licenseResponse.value ?? []).filter(
          (sku: {
            prepaidUnits?: { enabled?: number };
            skuPartNumber?: string;
          }) => (sku.prepaidUnits?.enabled ?? 0) > 0 && (sku.skuPartNumber ?? "").toLowerCase().indexOf("trial") === -1
        );

        const options = activeLicenses.map(
          (sku: {
            skuId: string;
            skuPartNumber: string;
            prepaidUnits: { enabled: number };
            consumedUnits: number;
          }) => {
            const friendlyName = friendlyNameMap[sku.skuPartNumber] || sku.skuPartNumber.replace(/_/g, " ");
            const total = sku.prepaidUnits.enabled;
            const consumed = sku.consumedUnits;
            const available = total - consumed;
            const availableText = total > 0 ? `${available} of ${total} available` : "Unlimited available";

            return {
              key: sku.skuId,
              text: `${friendlyName} - ${availableText}`,
            };
          }
        );

        if (!isMounted) {
          return;
        }

        setVerifiedDomain(domain);
        setLicenseOptions(options);
      } catch (fetchError) {
        console.error("Failed to load tenant data:", fetchError);

        if (isMounted) {
          setError("Failed to load tenant info or licenses. Make sure the required Graph permissions are granted.");
        }
      } finally {
        if (isMounted) {
          setLoadingTenantData(false);
        }
      }
    };

    fetchTenantData().catch((fetchError) => {
      console.error("Unexpected tenant data error:", fetchError);
    });

    return () => {
      isMounted = false;
    };
  }, [graphClient]);

  React.useEffect(() => {
    const firstInitial = formData.firstName.trim().charAt(0);
    const lastName = formData.lastName.trim();

    if (!firstInitial || !lastName) {
      return;
    }

    const suggestion = `${firstInitial}${lastName}`.toLowerCase().replace(/[^a-z0-9]/g, "");

    setFormData((previousFormData) => {
      if (!previousFormData.username || previousFormData.username === suggestedUsernameRef.current) {
        return { ...previousFormData, username: suggestion };
      }

      return previousFormData;
    });

    suggestedUsernameRef.current = suggestion;
  }, [formData.firstName, formData.lastName]);

  React.useEffect(() => {
    const trimmedUsername = formData.username.trim().toLowerCase();

    if (!trimmedUsername) {
      setUsernameAvailable(null);
      return;
    }

    setUsernameAvailable(null);
    let isCancelled = false;

    const timeoutId = window.setTimeout(async () => {
      try {
        const upn = `${trimmedUsername}@${verifiedDomain}`;
        await graphClient.api(`/users/${upn}`).get();

        if (!isCancelled) {
          setUsernameAvailable(false);
        }
      } catch {
        if (!isCancelled) {
          setUsernameAvailable(true);
        }
      }
    }, 700);

    return () => {
      isCancelled = true;
      window.clearTimeout(timeoutId);
    };
  }, [formData.username, verifiedDomain, graphClient]);

  const handleChange = <K extends keyof IFormData>(field: K, value: IFormData[K]): void => {
    if (field === "password" && typeof value === "string") {
      setPasswordStrength(validatePassword(value));
    }

    setFormData((previousFormData) => ({
      ...previousFormData,
      [field]: value,
    }));
  };

  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const file = event.target.files && event.target.files[0] ? event.target.files[0] : undefined;

    if (file) {
      const reader = new FileReader();
      reader.onload = () => setPhotoPreview(reader.result as string);
      reader.readAsDataURL(file);
    } else {
      setPhotoPreview(null);
    }

    setFormData((previousFormData) => ({
      ...previousFormData,
      photo: file,
    }));
  };

  const resetForm = (): void => {
    setFormData(createInitialFormData());
    setPhotoPreview(null);
    setPasswordVisible(false);
    setPasswordStrength({ label: "", tone: "none" });
    setUsernameAvailable(null);
    suggestedUsernameRef.current = "";

    if (fileInputRef.current) {
      fileInputRef.current.value = "";
    }
  };

  const triggerFileSelect = (): void => {
    fileInputRef.current?.click();
  };

  const createUser = async (): Promise<void> => {
    const trimmedFirstName = formData.firstName.trim();
    const trimmedLastName = formData.lastName.trim();
    const trimmedUsername = formData.username.trim().toLowerCase();

    setError(null);
    setSuccessMessage(null);

    if (!trimmedFirstName || !trimmedLastName || !trimmedUsername || !formData.password) {
      setError("Please complete all required account fields before creating the user.");
      return;
    }

    if (usernameAvailable === null) {
      setError("Wait for the username availability check to finish before creating the user.");
      return;
    }

    if (usernameAvailable === false) {
      setError("Choose a different username. The current one is already in use.");
      return;
    }

    if (formData.password !== formData.confirmPassword) {
      setError("The password and confirm password fields must match.");
      return;
    }

    if (passwordStrength.tone !== "strong") {
      setError("Use a strong password with at least 8 characters, upper and lower case letters, a number, and a symbol.");
      return;
    }

    setLoading(true);

    try {
      const userPrincipalName = `${trimmedUsername}@${verifiedDomain}`;
      const userPayload: Record<string, unknown> = {
        accountEnabled: true,
        displayName: `${trimmedFirstName} ${trimmedLastName}`,
        givenName: trimmedFirstName,
        surname: trimmedLastName,
        mailNickname: trimmedUsername,
        userPrincipalName,
        usageLocation: "US",
        passwordProfile: {
          forceChangePasswordNextSignIn: formData.forcePasswordChange,
          password: formData.password,
        },
      };

      if (formData.jobTitle.trim()) {
        userPayload.jobTitle = formData.jobTitle.trim();
      }

      if (formData.company.trim()) {
        userPayload.companyName = formData.company.trim();
      }

      if (formData.businessPhone.trim()) {
        userPayload.businessPhones = [formData.businessPhone.trim()];
      }

      if (formData.mobilePhone.trim()) {
        userPayload.mobilePhone = formData.mobilePhone.trim();
      }

      if (formData.fax.trim()) {
        userPayload.faxNumber = formData.fax.trim();
      }

      if (formData.address.trim()) {
        userPayload.streetAddress = formData.address.trim();
      }

      if (formData.city.trim()) {
        userPayload.city = formData.city.trim();
      }

      const user = await graphClient.api("/users").post(userPayload);

      if (formData.license) {
        try {
          await graphClient.api(`/users/${user.id}/assignLicense`).post({
            addLicenses: [{ skuId: formData.license }],
            removeLicenses: [],
          });
        } catch (licenseError) {
          console.error("License assignment failed:", licenseError);
        }
      }

      if (formData.photo) {
        const arrayBuffer = await formData.photo.arrayBuffer();
        const blob = new Blob([arrayBuffer], { type: formData.photo.type || "image/jpeg" });

        for (let attempt = 1; attempt <= 3; attempt += 1) {
          try {
            await new Promise((resolve) => {
              window.setTimeout(resolve, attempt * 5000);
            });

            await graphClient.api(`/users/${user.id}/photo/$value`).header("Content-Type", blob.type).put(blob);
            break;
          } catch (uploadError) {
            console.warn(`Photo upload attempt ${attempt} failed:`, uploadError);

            if (attempt === 3) {
              console.error("Photo upload failed after 3 retries.");
            }
          }
        }
      }

      setSuccessMessage(`User created successfully: ${userPrincipalName}`);
      window.setTimeout(() => setSuccessMessage(null), 6000);
      resetForm();
    } catch (createError) {
      console.error("Error creating user:", createError);
      setError("The account could not be created. Check the console for more details and confirm the Graph permissions are correct.");
    } finally {
      setLoading(false);
    }
  };

  const initials = `${formData.firstName.trim().charAt(0).toUpperCase()}${formData.lastName.trim().charAt(0).toUpperCase()}` || "AU";
  const previewName = [formData.firstName.trim(), formData.lastName.trim()].filter(Boolean).join(" ") || "New user preview";
  const previewEmail = formData.username.trim() ? `${formData.username.trim().toLowerCase()}@${verifiedDomain}` : `username@${verifiedDomain}`;
  const selectedLicenseLabel =
    licenseOptions.filter((option: IDropdownOption) => option.key === formData.license)[0]?.text || "No license selected";
  const completionChecks = [
    Boolean(formData.firstName.trim()),
    Boolean(formData.lastName.trim()),
    Boolean(formData.username.trim()),
    Boolean(formData.password),
    Boolean(formData.confirmPassword),
    Boolean(formData.company.trim()),
    Boolean(formData.jobTitle.trim()),
    Boolean(formData.businessPhone.trim()),
    Boolean(formData.city.trim()),
    Boolean(formData.license),
    Boolean(formData.photo),
  ];
  const completionCount = completionChecks.filter(Boolean).length;
  const completionPercent = completionCount / completionChecks.length;
  const passwordToneClass =
    passwordStrength.tone === "strong"
      ? styles.statusSuccess
      : passwordStrength.tone === "medium"
      ? styles.statusWarning
      : passwordStrength.tone === "weak"
      ? styles.statusDanger
      : styles.statusNeutral;
  const usernameStatusClass =
    !formData.username.trim() || usernameAvailable === null
      ? styles.statusNeutral
      : usernameAvailable
      ? styles.statusSuccess
      : styles.statusDanger;
  const usernameStatusText = !formData.username.trim()
    ? "Enter a username to verify availability."
    : usernameAvailable === null
    ? "Checking username availability..."
    : usernameAvailable
    ? "Username is available."
    : "Username already exists.";
  const createDisabled =
    loading ||
    loadingTenantData ||
    (!formData.username.trim() ? false : usernameAvailable === null || usernameAvailable === false);

  return (
    <div className={styles.shell}>
      <div className={styles.surface}>
        <div className={styles.hero}>
          <div className={styles.heroMain}>
            <div className={styles.heroBadge}>
              <Icon iconName="AddFriend" className={styles.heroBadgeIcon} />
              <span>Modern Entra ID onboarding workspace</span>
            </div>

            <div className={styles.headerRow}>
              <div className={styles.headerContent}>
                <h2 className={styles.header}>Create New AD User</h2>
                <p className={styles.subheader}>
                  Provision the account, apply the right license, and prepare the profile in one focused flow.
                </p>
              </div>

              {onClose && <DefaultButton text="Close" onClick={onClose} />}
            </div>

            <ul className={styles.heroHighlights}>
              <li className={styles.heroHighlightItem}>
                <span className={styles.heroHighlightBullet} aria-hidden="true" />
                <div className={styles.heroHighlightText}>
                  <div className={styles.heroHighlightLabel}>Default domain</div>
                  <div className={styles.heroHighlightValue}>@{verifiedDomain}</div>
                </div>
              </li>
              <li className={styles.heroHighlightItem}>
                <span className={styles.heroHighlightBullet} aria-hidden="true" />
                <div className={styles.heroHighlightText}>
                  <div className={styles.heroHighlightLabel}>License catalog</div>
                  <div className={styles.heroHighlightValue}>
                    {loadingTenantData ? "Syncing..." : `${licenseOptions.length} active SKUs`}
                  </div>
                </div>
              </li>
              <li className={styles.heroHighlightItem}>
                <span className={styles.heroHighlightBullet} aria-hidden="true" />
                <div className={styles.heroHighlightText}>
                  <div className={styles.heroHighlightLabel}>Password policy</div>
                  <div className={styles.heroHighlightValue}>Strong password required</div>
                </div>
              </li>
            </ul>
          </div>

          <div className={styles.previewCard}>
            <div className={styles.previewHeader}>
              <div className={styles.previewEyebrow}>Account preview</div>
              <div className={styles.previewCompletion}>{Math.round(completionPercent * 100)}% complete</div>
            </div>

            <div className={styles.previewProfile}>
              <div className={styles.personaWrapper} onClick={triggerFileSelect}>
                <Persona
                  imageUrl={photoPreview || undefined}
                  imageInitials={initials}
                  size={PersonaSize.size100}
                  hidePersonaDetails
                />
                <div className={styles.cameraOverlay}>
                  <Icon iconName="Camera" className={styles.cameraIcon} />
                </div>
              </div>

              <div className={styles.previewIdentity}>
                <div className={styles.previewName}>{previewName}</div>
                <div className={styles.previewEmail}>{previewEmail}</div>
              </div>
            </div>

            <DefaultButton
              className={styles.photoButton}
              text={photoPreview ? "Change photo" : "Upload photo"}
              onClick={triggerFileSelect}
            />
            <input
              ref={fileInputRef}
              type="file"
              accept="image/*"
              onChange={handleFileChange}
              className={styles.hiddenInput}
            />

            <div className={styles.progressBlock}>
              <ProgressIndicator
                label={`${completionCount} of ${completionChecks.length} key details ready`}
                percentComplete={completionPercent}
              />
            </div>

            <div className={styles.previewMetaGrid}>
              <div className={styles.previewMetaItem}>
                <div className={styles.previewMetaLabel}>Username</div>
                <div className={styles.previewMetaValue}>{formData.username.trim() || "Pending"}</div>
              </div>
              <div className={styles.previewMetaItem}>
                <div className={styles.previewMetaLabel}>Role</div>
                <div className={styles.previewMetaValue}>{formData.jobTitle.trim() || "Not set"}</div>
              </div>
              <div className={styles.previewMetaItem}>
                <div className={styles.previewMetaLabel}>Company</div>
                <div className={styles.previewMetaValue}>{formData.company.trim() || "Not set"}</div>
              </div>
              <div className={styles.previewMetaItem}>
                <div className={styles.previewMetaLabel}>License</div>
                <div className={styles.previewMetaValue}>{selectedLicenseLabel}</div>
              </div>
            </div>
          </div>
        </div>

        <div className={styles.messageStack}>
          {error && (
            <MessageBar messageBarType={MessageBarType.error} isMultiline={false} onDismiss={() => setError(null)}>
              {error}
            </MessageBar>
          )}
          {successMessage && (
            <MessageBar
              messageBarType={MessageBarType.success}
              isMultiline={false}
              onDismiss={() => setSuccessMessage(null)}
            >
              {successMessage}
            </MessageBar>
          )}
        </div>

        <div className={styles.sectionGrid}>
          <section className={`${styles.sectionCard} ${styles.sectionCardWide}`}>
            {renderSectionHeading(
              "Contact",
              "Profile",
              "Identity details",
              "Capture the basics for the account display name, directory identity, and organization."
            )}

            <div className={styles.formGrid}>
              <TextField
                label="First name"
                required
                value={formData.firstName}
                onChange={(_, value) => handleChange("firstName", value || "")}
              />
              <TextField
                label="Last name"
                required
                value={formData.lastName}
                onChange={(_, value) => handleChange("lastName", value || "")}
              />
              <TextField
                label="Company"
                value={formData.company}
                onChange={(_, value) => handleChange("company", value || "")}
              />
              <TextField
                label="Job title"
                value={formData.jobTitle}
                onChange={(_, value) => handleChange("jobTitle", value || "")}
              />

              <div className={`${styles.fieldBlock} ${styles.fullSpan}`}>
                <TextField
                  label="Username"
                  required
                  value={formData.username}
                  suffix={`@${verifiedDomain}`}
                  onChange={(_, value) => handleChange("username", (value || "").replace(/\s+/g, ""))}
                />

                <div className={styles.inlineStatusRow}>
                  <div className={`${styles.statusPill} ${usernameStatusClass}`}>{usernameStatusText}</div>
                  <div className={styles.helperText}>This becomes the primary sign-in ID for the new user.</div>
                </div>
              </div>
            </div>
          </section>

          <section className={styles.sectionCard}>
            {renderSectionHeading(
              "Shield",
              "Security",
              "Password setup",
              "Give the account a secure starting password and control first sign-in behavior."
            )}

            <div className={styles.formGrid}>
              <div className={styles.passwordField}>
                <TextField
                  label="Password"
                  type={passwordVisible ? "text" : "password"}
                  required
                  value={formData.password}
                  onChange={(_, value) => handleChange("password", value || "")}
                />
                <IconButton
                  iconProps={{ iconName: passwordVisible ? "Hide" : "RedEye" }}
                  ariaLabel={passwordVisible ? "Hide password" : "Show password"}
                  onClick={() => setPasswordVisible((previousValue) => !previousValue)}
                  className={styles.eyeIcon}
                />
              </div>
              <TextField
                label="Confirm password"
                type="password"
                required
                value={formData.confirmPassword}
                onChange={(_, value) => handleChange("confirmPassword", value || "")}
              />

              <div className={`${styles.fieldBlock} ${styles.fullSpan}`}>
                <div className={styles.inlineStatusRow}>
                  <div className={`${styles.statusPill} ${passwordToneClass}`}>
                    {passwordStrength.label || "Use a strong password to continue."}
                  </div>
                  <div className={styles.helperText}>
                    Include at least 8 characters with upper and lower case letters, a number, and a symbol.
                  </div>
                </div>
              </div>
            </div>

            <div className={styles.toggleRow}>
              <Toggle
                label="Require password change on next sign-in"
                checked={formData.forcePasswordChange}
                onChange={(_, checked) => handleChange("forcePasswordChange", !!checked)}
              />
              <div className={styles.helperText}>Recommended when handing credentials to a new starter.</div>
            </div>
          </section>

          <section className={styles.sectionCard}>
            {renderSectionHeading(
              "Phone",
              "Contact",
              "Reachability",
              "Add contact channels that help teams and admins reach the user quickly."
            )}

            <div className={styles.formGrid}>
              <TextField
                label="Business phone"
                value={formData.businessPhone}
                onChange={(_, value) => handleChange("businessPhone", value || "")}
              />
              <TextField
                label="Mobile phone"
                value={formData.mobilePhone}
                onChange={(_, value) => handleChange("mobilePhone", value || "")}
              />
              <div className={`${styles.fieldBlock} ${styles.fullSpan}`}>
                <TextField label="Fax" value={formData.fax} onChange={(_, value) => handleChange("fax", value || "")} />
              </div>
            </div>
          </section>

          <section className={`${styles.sectionCard} ${styles.sectionCardWide}`}>
            {renderSectionHeading(
              "MapPin",
              "Location & access",
              "Address and licensing",
              "Set optional location details and assign the right subscription at provisioning time."
            )}

            <div className={styles.formGrid}>
              <div className={`${styles.fieldBlock} ${styles.fullSpan}`}>
                <TextField
                  label="Street address"
                  multiline
                  rows={2}
                  value={formData.address}
                  onChange={(_, value) => handleChange("address", value || "")}
                />
              </div>
              <TextField label="City" value={formData.city} onChange={(_, value) => handleChange("city", value || "")} />

              <div className={styles.fieldBlock}>
                <Dropdown
                  label="License package"
                  placeholder={loadingTenantData ? "Loading licenses..." : "Select a license"}
                  options={licenseOptions}
                  selectedKey={formData.license || undefined}
                  onChange={(_, option) => handleChange("license", (option?.key as string) || "")}
                  disabled={loadingTenantData || licenseOptions.length === 0}
                />
              </div>

              <div className={`${styles.fieldBlock} ${styles.fullSpan}`}>
                <div className={styles.inlineStatusRow}>
                  <div className={`${styles.statusPill} ${formData.license ? styles.statusSuccess : styles.statusNeutral}`}>
                    {loadingTenantData ? "Loading available licenses..." : selectedLicenseLabel}
                  </div>
                  <div className={styles.helperText}>Only active, non-trial tenant subscriptions are shown.</div>
                </div>
              </div>
            </div>
          </section>
        </div>

        <div className={styles.actionBar}>
          <div className={styles.actionHint}>
            {loadingTenantData
              ? "Syncing your tenant configuration before provisioning is available."
              : `New accounts will be created in ${verifiedDomain}.`}
          </div>

          <div className={styles.buttonRow}>
            <DefaultButton text="Reset form" onClick={resetForm} disabled={loading} />
            <PrimaryButton text="Create user" onClick={createUser} disabled={createDisabled} />
          </div>
        </div>

        {loading && (
          <div className={styles.loadingOverlay}>
            <div className={styles.loadingCard}>
              <Spinner size={SpinnerSize.large} label="Creating account and applying the selected setup..." />
            </div>
          </div>
        )}
      </div>
    </div>
  );
};

export default AdUserForm;
