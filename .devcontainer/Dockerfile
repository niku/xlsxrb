FROM ruby:4.0

# Install .NET SDK for E2E testing with Open-XML-SDK.
# Use dotnet-install so the container works on both arm64 and amd64 hosts.
RUN apt-get update && apt-get install -y curl ca-certificates && \
    curl -sSL https://dot.net/v1/dotnet-install.sh -o /tmp/dotnet-install.sh && \
    bash /tmp/dotnet-install.sh --channel 8.0 --quality ga --install-dir /usr/share/dotnet && \
    ln -s /usr/share/dotnet/dotnet /usr/bin/dotnet && \
    rm -f /tmp/dotnet-install.sh && \
    apt-get clean && rm -rf /var/lib/apt/lists/*

# Pre-download DocumentFormat.OpenXml package to the NuGet cache for faster tool builds later
RUN mkdir /tmp/nuget_warmup && cd /tmp/nuget_warmup && \
    dotnet new console && \
    dotnet add package DocumentFormat.OpenXml --version 3.5.1 && \
    cd / && rm -rf /tmp/nuget_warmup

# Set up the workspace
WORKDIR /workspace

# Install gems (Gemfile.lock will be respected)
COPY Gemfile xlsxrb.gemspec ./
COPY lib/xlsxrb/version.rb ./lib/xlsxrb/
RUN bundle install

# Copy the rest of the application
COPY . .
