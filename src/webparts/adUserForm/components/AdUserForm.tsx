import * as React from "react";
import {
  TextField,
  Dropdown,
  IDropdownOption,
  Toggle,
  PrimaryButton,
  DefaultButton,
  Label,
  Separator,
  Spinner,
  SpinnerSize,
  Persona,
  PersonaSize,
  IconButton,
  Icon,
  Stack,
} from "@fluentui/react";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import styles from "./AdUserForm.module.scss";

export interface IAdUserFormProps {
  graphClient: MSGraphClientV3;
  onClose?: () => void;
}

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

const AdUserForm: React.FC<IAdUserFormProps> = ({ graphClient, onClose }) => {
  const [formData, setFormData] = React.useState({
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
    photo: null as File | null,
  });

  const [photoPreview, setPhotoPreview] = React.useState<string | null>(null);
  const [licenseOptions, setLicenseOptions] = React.useState<IDropdownOption[]>([]);
  const [verifiedDomain, setVerifiedDomain] = React.useState("sj30b.onmicrosoft.com");
  const [loading, setLoading] = React.useState(false);
  const [error, setError] = React.useState<string | null>(null);
  const [passwordVisible, setPasswordVisible] = React.useState(false);
  const [passwordStrength, setPasswordStrength] = React.useState<{ label: string; color: string }>({ label: "", color: "" });
  const [usernameAvailable, setUsernameAvailable] = React.useState<null | boolean>(null);
  const fileInputRef = React.useRef<HTMLInputElement>(null);

  // Fetch tenant info + licenses
  React.useEffect(() => {
    const fetchLicensesAndDomain = async () => {
      try {
        const orgResponse = await graphClient.api("/organization").get();
        const domain =
          orgResponse.value[0].verifiedDomains.find((d: any) => d.isDefault)?.name ||
          "sj30b.onmicrosoft.com";
        setVerifiedDomain(domain);

        const licenseResponse = await graphClient.api("/subscribedSkus").get();
        const activeLicenses = licenseResponse.value.filter(
          (sku: any) => sku.prepaidUnits.enabled > 0 && !sku.skuPartNumber.toLowerCase().includes("trial")
        );

        const options = activeLicenses.map((sku: any) => {
          const friendlyName =
            friendlyNameMap[sku.skuPartNumber] ||
            sku.skuPartNumber.replace(/_/g, " ");
          const total = sku.prepaidUnits.enabled;
          const consumed = sku.consumedUnits;
          const available = total - consumed;
          const availableText =
            total > 0
              ? `${available} of ${total} available`
              : "Unlimited available";

          return {
            key: sku.skuId,
            text: `${friendlyName} — ${availableText}`,
          };
        });

        setLicenseOptions(options);
      } catch (err) {
        console.error("Fetch failed:", err);
        setError("Failed to load licenses or tenant info. Ensure Graph API permissions are granted.");
      }
    };
    fetchLicensesAndDomain();
  }, [graphClient]);

  // Suggest username
  React.useEffect(() => {
    if (formData.firstName && formData.lastName) {
      const suggestion = (formData.firstName[0] + formData.lastName)
        .toLowerCase()
        .replace(/[^a-z0-9]/g, "");
      handleChange("username", suggestion);
    }
  }, [formData.firstName, formData.lastName]);

  // Runtime check: username availability
  React.useEffect(() => {
    const checkUser = async () => {
      if (!formData.username) return;
      try {
        const upn = `${formData.username}@${verifiedDomain}`;
        await graphClient.api(`/users/${upn}`).get();
        setUsernameAvailable(false);
      } catch {
        setUsernameAvailable(true);
      }
    };
    const delay = setTimeout(checkUser, 800);
    return () => clearTimeout(delay);
  }, [formData.username, verifiedDomain]);

  // Password validation
  const validatePassword = (password: string) => {
    const strongRegex = /^(?=.*[a-z])(?=.*[A-Z])(?=.*\d)(?=.*[^A-Za-z\d]).{8,}$/;
    const mediumRegex = /^(?=.*[A-Z])(?=.*\d).{6,}$/;
    if (!password) return { label: "", color: "" };
    if (strongRegex.test(password)) return { label: "Strong password", color: "green" };
    if (mediumRegex.test(password)) return { label: "Moderate password", color: "orange" };
    return { label: "Weak password", color: "red" };
  };

  const handleChange = (field: string, value: string | boolean) => {
    if (field === "password") setPasswordStrength(validatePassword(value as string));
    setFormData({ ...formData, [field]: value });
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0] || null;
    if (file) {
      const reader = new FileReader();
      reader.onload = () => setPhotoPreview(reader.result as string);
      reader.readAsDataURL(file);
    }
    setFormData({ ...formData, photo: file });
  };

  const triggerFileSelect = () => fileInputRef.current?.click();

  // Create user
  const createUser = async () => {
    setError(null);
    if (!formData.firstName || !formData.lastName || !formData.username || !formData.password) {
      setError("Please fill all required fields.");
      return;
    }
    if (formData.password !== formData.confirmPassword) {
      setError("Passwords do not match.");
      return;
    }
    if (passwordStrength.color !== "green") {
      setError("Please choose a strong password before creating user.");
      return;
    }

    setLoading(true);
    try {
      const mailNickname = formData.username.toLowerCase();
      const userPrincipalName = `${mailNickname}@${verifiedDomain}`;

      const userPayload = {
        accountEnabled: true,
        displayName: `${formData.firstName} ${formData.lastName}`,
        givenName: formData.firstName,
        surname: formData.lastName,
        mailNickname,
        userPrincipalName,
        jobTitle: formData.jobTitle,
        companyName: formData.company,
        usageLocation: "US",
        passwordProfile: {
          forceChangePasswordNextSignIn: formData.forcePasswordChange,
          password: formData.password,
        },
        businessPhones: [formData.businessPhone],
        mobilePhone: formData.mobilePhone,
        faxNumber: formData.fax,
        streetAddress: formData.address,
        city: formData.city,
      };

      const user = await graphClient.api("/users").post(userPayload);
      console.log("✅ User created:", user);

      if (formData.photo) {
        const arrayBuffer = await formData.photo.arrayBuffer();
        const blob = new Blob([arrayBuffer], { type: formData.photo.type || "image/jpeg" });
        await graphClient
          .api(`/users/${user.id}/photo/$value`)
          .header("Content-Type", blob.type)
          .put(blob);
      }

      if (formData.license) {
        await graphClient.api(`/users/${user.id}/assignLicense`).post({
          addLicenses: [{ skuId: formData.license }],
          removeLicenses: [],
        });
      }

      alert(`✅ User created successfully: ${userPrincipalName}`);
    } catch (err: any) {
      console.error("❌ Error creating user:", err);
      setError("Error creating user. Check console or permissions.");
    } finally {
      setLoading(false);
    }
  };

  const initials = `${(formData.firstName[0] || "").toUpperCase()}${(formData.lastName[0] || "").toUpperCase()}`;

  return (
    <div className={styles.card}>
      <div className={styles.headerRow}>
        <h2 className={styles.header}>Create New AD User</h2>
        {onClose && <DefaultButton text="Close" onClick={onClose} />}
      </div>

      {error && <div className={styles.errorBox}>❌ {error}</div>}

      {/* Profile photo */}
      <Stack verticalAlign="center" horizontalAlign="center" className={styles.avatarSection}>
        <div className={styles.personaWrapper} onClick={triggerFileSelect}>
          <Persona
            text=""
            imageUrl={photoPreview || undefined}
            imageInitials={initials || "FL"}
            size={PersonaSize.size100}
            hidePersonaDetails={true}
          />
          <div className={styles.cameraOverlay}>
            <Icon iconName="Camera" className={styles.cameraIcon} />
          </div>
        </div>
        <Label className={styles.photoLabel}>Profile Photo</Label>
        <input type="file" accept="image/*" onChange={handleFileChange} ref={fileInputRef} className={styles.hiddenInput} />
      </Stack>

      <Separator />

      {/* Basic Info */}
      <div className={styles.formGrid}>
        <TextField label="First Name" required onChange={(_, v) => handleChange("firstName", v || "")} />
        <TextField label="Last Name" required onChange={(_, v) => handleChange("lastName", v || "")} />
        <TextField label="Company" onChange={(_, v) => handleChange("company", v || "")} />
        <TextField label="Job Title" onChange={(_, v) => handleChange("jobTitle", v || "")} />
        <TextField
          label="Username"
          required
          value={formData.username}
          onChange={(_, v) => handleChange("username", v || "")}
          iconProps={
            usernameAvailable === null
              ? undefined
              : usernameAvailable
                ? { iconName: "CheckMark", style: { color: "green" } }
                : { iconName: "StatusErrorFull", style: { color: "red" } }
          }
        />
        <TextField label="Domain" readOnly value={`@${verifiedDomain}`} />

        <div className={styles.passwordField}>
          <TextField
            label="Password"
            type={passwordVisible ? "text" : "password"}
            required
            onChange={(_, v) => handleChange("password", v || "")}
            canRevealPassword={false}
          />
          <IconButton
            iconProps={{ iconName: passwordVisible ? "Hide" : "RedEye" }}
            onClick={() => setPasswordVisible(!passwordVisible)}
            className={styles.eyeIcon}
          />
          {passwordStrength.label && (
            <div style={{ color: passwordStrength.color, fontSize: "12px" }}>
              {passwordStrength.label}
            </div>
          )}
        </div>


        <TextField
          label="Confirm Password"
          type="password"
          required
          onChange={(_, v) => handleChange("confirmPassword", v || "")}
        />
      </div>

      <Toggle
        label="Require password change on next sign-in"
        checked={formData.forcePasswordChange}
        onChange={(_, v) => handleChange("forcePasswordChange", !!v)}
      />

      <Separator />
      <Label className={styles.sectionLabel}>Contact Information</Label>
      <div className={styles.formGrid}>
        <TextField label="Business Phone" onChange={(_, v) => handleChange("businessPhone", v || "")} />
        <TextField label="Mobile Phone" onChange={(_, v) => handleChange("mobilePhone", v || "")} />
        <TextField label="Fax" onChange={(_, v) => handleChange("fax", v || "")} />
      </div>

      <Separator />
      <Label className={styles.sectionLabel}>Address</Label>
      <div className={styles.formGrid}>
        <TextField label="Street Address" onChange={(_, v) => handleChange("address", v || "")} />
        <TextField label="City" onChange={(_, v) => handleChange("city", v || "")} />
      </div>

      <Separator />
      <Label className={styles.sectionLabel}>Assign License</Label>
      <div className={styles.licenseRow}>
        <Dropdown
          placeholder="Select a license"
          options={licenseOptions}
          onChange={(_, option) => handleChange("license", option?.key as string)}
          disabled={licenseOptions.length === 0}
        />
      </div>

      <Stack horizontal horizontalAlign="end" tokens={{ childrenGap: 12 }} className={styles.buttonRow}>
        <DefaultButton text="Cancel" onClick={onClose} />
        <PrimaryButton text="Create" onClick={createUser} disabled={loading} />
      </Stack>

      {loading && (
        <div className={styles.spinnerContainer}>
          <Spinner size={SpinnerSize.large} label="Creating AD user..." />
        </div>
      )}
    </div>
  );
};

export default AdUserForm;
