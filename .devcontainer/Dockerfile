FROM ruby:4.0

ARG USERNAME=vscode
ARG USER_UID=1000
ARG USER_GID=$USER_UID

# Create a non-root user (vscode) with passwordless sudo privileges
# following Microsoft's Dev Container guidelines to prevent file permission conflicts.
RUN apt-get update && apt-get install -y sudo ca-certificates \
    # Check if GID already exists, otherwise create it
    && if [ "$(getent group ${USER_GID})" ]; then \
    echo "GID ${USER_GID} already exists."; \
    else \
    groupadd --gid ${USER_GID} ${USERNAME}; \
    fi \
    # Check if UID already exists, otherwise create user
    && if [ "$(getent passwd ${USER_UID})" ]; then \
    echo "UID ${USER_UID} already exists."; \
    else \
    useradd --uid ${USER_UID} --gid ${USER_GID} -m ${USERNAME} -s /bin/bash; \
    fi \
    # Grant passwordless sudo permissions
    && echo "${USERNAME} ALL=(root) NOPASSWD:ALL" > /etc/sudoers.d/${USERNAME} \
    && chmod 0440 /etc/sudoers.d/${USERNAME} \
    && apt-get clean && rm -rf /var/lib/apt/lists/*

# Install .NET SDK for E2E testing with Open-XML-SDK.
# Use dotnet-install so the container works on both arm64 and amd64 hosts.
RUN apt-get update && apt-get install -y curl ca-certificates && \
    curl -sSL https://dot.net/v1/dotnet-install.sh -o /tmp/dotnet-install.sh && \
    bash /tmp/dotnet-install.sh --channel 8.0 --quality ga --install-dir /usr/share/dotnet && \
    ln -s /usr/share/dotnet/dotnet /usr/bin/dotnet && \
    rm -f /tmp/dotnet-install.sh && \
    apt-get clean && rm -rf /var/lib/apt/lists/*

# Install LibreOffice, poppler-utils, ImageMagick, and fonts for Visual Examples/VRT
RUN apt-get update && apt-get install -y \
    libreoffice-calc \
    poppler-utils \
    imagemagick \
    fonts-liberation \
    fonts-noto-cjk \
    && apt-get clean && rm -rf /var/lib/apt/lists/*

# Pre-create the workspace folder and assign ownership to the vscode user
RUN mkdir -p /workspaces/xlsxrb && chown -R ${USERNAME}:${USERNAME} /workspaces

# Set up the workspace
USER ${USERNAME}
WORKDIR /workspaces/xlsxrb

# Pre-download DocumentFormat.OpenXml package to the NuGet cache
RUN mkdir /tmp/nuget_warmup && cd /tmp/nuget_warmup && \
    dotnet new console && \
    dotnet add package DocumentFormat.OpenXml --version 3.5.1 && \
    cd / && rm -rf /tmp/nuget_warmup

# Install gems (Respect Gemfile.lock, maintain correct user ownership)
COPY --chown=vscode:vscode Gemfile xlsxrb.gemspec ./
COPY --chown=vscode:vscode lib/xlsxrb/version.rb ./lib/xlsxrb/
RUN bundle install

# Copy the remaining project application files
COPY --chown=vscode:vscode . .
